<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �ٹ����� ����Ʈ ������
' History : 2011.03.10 ������ ����
'           2012.02.14 ������ - �̴ϴ޷� ��ü
'           2018.03.26 ������ - ���� ����Ʈ �߰�
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<!-- #include virtual="/lib/classes/member/tenAgitCalendarCls.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenDepartmentCls.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenAgitCls.asp" -->
<%
  dim empno, cMember
	dim idx, sDt, sTm, eDt, eTm
	dim AreaDiv,userid,susername,iposit_sn,ipart_sn,spart_name, sposit_name, sdepartmentNameFull
	dim suserPhone,susercell,ChkStart,ChkEnd,usePersonNo,etcComment
	dim department_id, inc_subdepartment, nodepartonly
	dim usepoint, usemoney
	dim nowDateShort, availSdate, availEdate, nowYearShort
	
	'// ��û ���� �Ⱓ ����
	nowDateShort = mid(date(),6,10)
	nowYearShort = year(date)
	if nowDateShort>="02-15" and nowDateShort<="04-14" then		'1�� ��û
		availSdate = dateserial(nowYearShort,2,17)
		availEdate = dateserial(nowYearShort,4,30)
	elseif nowDateShort>="04-15" and nowDateShort<="06-14" then	'2�� ��û
		availSdate = dateserial(nowYearShort,4,17)
		availEdate = dateserial(nowYearShort,6,30)
	elseif nowDateShort>="06-15" and nowDateShort<="08-14" then	'3�� ��û
		availSdate = dateserial(nowYearShort,6,15)
		availEdate = dateserial(nowYearShort,8,31)
	elseif nowDateShort>="08-15" and nowDateShort<="10-14" then	'4�� ��û
		availSdate = dateserial(nowYearShort,8,17)
		availEdate = dateserial(nowYearShort,10,31)
	elseif nowDateShort>="10-15" and nowDateShort<="12-14" then	'5�� ��û
		availSdate = dateserial(nowYearShort,10,17)
		availEdate = dateserial(nowYearShort,12,31)
	elseif nowDateShort>="12-15" or nowDateShort<="02-14" then	'6�� ��û(�ͳ�)
		availSdate = dateserial(chkIIF(nowDateShort<="02-14",nowYearShort-1,nowYearShort),12,12)
		availEdate = dateserial(chkIIF(nowDateShort<="02-14",nowYearShort,nowYearShort+1),2,28)
	end if

	idx = request("idx")
	empno = session("ssBctSn")
	sDt		= requestCheckvar(Request("ChkStart"),10)
 
 '// ���� ����
	if idx<>"" then
		dim oAgitCal
		Set oAgitCal = new CAgitCalendarDetail
		oAgitCal.read(idx)

		AreaDiv			= oAgitCal.FAreaDiv
		empno				= oAgitCal.Fempno
		userid			= oAgitCal.Fuserid
		susername		= oAgitCal.Fusername
		iposit_sn		= oAgitCal.Fposit_sn
		ipart_sn			= oAgitCal.Fpart_sn
		department_id =  oAgitCal.Fdepartment_id
		suserPhone		= oAgitCal.FuserPhone
		susercell			= oAgitCal.FuserHP
		ChkStart		= oAgitCal.FChkStart
		ChkEnd			= oAgitCal.FChkEnd
		usePersonNo		= oAgitCal.FusePersonNo
		etcComment		= oAgitCal.FetcComment
    usepoint      = oAgitCal.FusePoint
    usemoney      = oAgitCal.FuseMoney
    
		sDt = left(ChkStart,10)
		eDt = left(ChkEnd,10)
		sTm = Num2Str(Hour(ChkStart),2,"0","R") & ":" & Num2Str(Minute(ChkStart),2,"0","R")
		eTm = Num2Str(Hour(ChkEnd),2,"0","R") & ":" & Num2Str(Minute(ChkEnd),2,"0","R")

		Set oAgitCal = Nothing
	end if
	
 if empno <> "" then
Set cMember = new CTenByTenMember
	cMember.Fempno = empno
	cMember.fnGetMemberData
	
	empno   		= cMember.Fempno
	userid			= cMember.Fuserid 
	susername      	= cMember.Fusername 
	suserphone     	= cMember.Fuserphone
	susercell      	= cMember.Fusercell   
	ipart_sn       	= cMember.Fpart_sn
	iposit_sn     	= cMember.Fposit_sn 
	spart_name     	= cMember.Fpart_name
	sposit_name     = cMember.Fposit_name 
	sdepartmentNameFull	= cMember.FdepartmentNameFull
Set cMember = nothing

dim clsap
dim totap, useap, reqap,payap, ispenalty, psdate, pedate, pkind
set clsap = new CMyAgit
		clsap.FRectEmpno = empno
		clsap.FRectChkStart = sDt
		clsap.fnGetMyAgit
		totap = clsap.FtotPoint
		useap = clsap.FusePoint 
		pkind = clsap.Fpenaltykind
		psdate = clsap.Fpenaltysdate
		pedate = clsap.Fpenaltyedate
set clsap = nothing
end if

	

	''if AreaDiv="" then AreaDiv="1"
	if sDt="" then sDt=date
	'if sTm="" then
	 sTm="15:00" 
	if eDt="" then eDt=dateAdd("d",1,sDt)
'	if eTm="" then
		 eTm="13:00"
	if usePersonNo="" then usePersonNo=1
		
		dim chkdate
		chkdate = datediff("d",date(),sDt)
	 
%>
<script language="javascript1.2" type="text/javascript" src="/js/datetime.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language="javascript">
<!--
	// ��� Ȯ�� �� ó��
	function jsSubmit()	{
		var frm = document.frm;
		if(frm.AreaDiv.value == "") {
			alert("������ ����Ʈ�� �������ּ���.");
			frm.AreaDiv.focus();
			return ;
		}  
		
		if(frm.userPhone.value==""&&frm.userHP.value=="") {
			alert("��󿬶�ó�� �Է����ּ���.");
			return;
		}

		if(getDayInterval(toDate(frm.ChkStart.value), toDate('<%=date%>'))>0) {
			alert("������ ��¥�� ����Ͻ� �� �����ϴ�. ��¥�� Ȯ�����ּ���.");
			return ;
		}

		if(getDayInterval(toDate(frm.ChkStart.value), toDate(frm.ChkEnd.value))<0) {
			alert("�Ⱓ�� �߸��Ǿ� �ֽ��ϴ�. ��¥�� Ȯ�����ּ���.");
			return ;
		}
		
		//�Է°��� �Ⱓ üũ
		if(frm.ChkStart.value < "<%=availSdate%>") {
			alert("���� ���� ��¥�� �ƴմϴ�.\n�� <%=availSdate%>���� ����");
			return;
		}
		if(frm.ChkEnd.value > "<%=availEdate%>") {
			alert("���� ���� ��¥�� �ƴմϴ�.\n�� <%=availEdate%>���� ����");
			return;
		}

		if(frm.uTerm.value>5){
			alert("�̿�Ⱓ�� �ִ� 5�ڱ��� �����մϴ�.");
			return;
		}

		if(frm.chkp.value!=1){
			alert("����Ʈ/ �ݾ�Ȯ�� ��ư�� �����ּ���");
			return;
		}
		
		if(frm.iPoint.value==""){
			alert("����Ʈ�� Ȯ�����ּ���");
			return;
		}
		
			if(frm.sMoney.value==""){
			alert("�ݾ��� Ȯ�����ּ���");
			return;
		}
		 
		if(parseInt(frm.iPoint.value)>parseInt(frm.avPoint.value)){
			alert("��뿹������Ʈ�� ��밡�� ����Ʈ���� �����ϴ�. ��û �Ұ����մϴ�.");
			return;
		}
		
		if(confirm(frm.ChkStart.value +"~"+frm.ChkEnd.value +"�Ⱓ�� ("+frm.uTerm.value+"��)   ����Ͻðڽ��ϱ�?"))	{
	 			document.frm.submit();
		}
		 
	}

	//�̿� �Ⱓ Ȯ�� �� �ڼ� �ڵ��Է�
	function chkTerm() {
		var frm = document.frm;
		
		var startday = frm.ChkStart;
		var endday = frm.ChkEnd;
	
		var startdate = toDate(startday.value);
		var enddate = toDate(endday.value);
	
		if ((startday.value == "") || (endday.value == "")) {
			alert("�Ⱓ�� �Է����ֽʽÿ�.");
			return;
		}
	
		if (getDayInterval(startdate, enddate) < 0) {
			//alert("�߸��� �Ⱓ�Դϴ�.");
			//return;
		}
	
		frm.uTerm.value = getDayInterval(startdate, enddate);
		frm.chkp.value = 0;
	}

	//���� ���̵� �˻� �� ���ó��� �ڵ��Է�
	function chkTenMember(uid) {
		if(!uid) {
			alert("�˻��� ID�� �Է����ּ���.");
			frm.userid.focus();
			return;
		}

		document.getElementById("ifmProc").src="actionTenUser.asp?uid="+uid;
	}

	//��üó��
	function delBook() {
		var chkdate = "<%=chkdate%>";
		var sMsg ="";
		if (chkdate <6  && chkdate >0){
			 //sMsg="������ 5����~1���� ��Ҵ� 3������ �̿��� �Ұ����ϰ� ��û ����Ʈ�� �����˴ϴ�."  
		}else if(chkdate ==0){
			//sMsg="������ ���� ��Ҵ� 6������ �̿��� �Ұ����ϰ� ��û ����Ʈ�� �����Ǹ� ȯ���� �Ұ��մϴ�."  
		}
		
		if(confirm(sMsg+" ��û�� ��� �Ͻðڽ��ϱ�?"))	{
			frm.mode.value = "del";
			frm.submit();
		}
	}
	
	//����Ʈ, �ݾ� ����
	function jsSetPoint(){
		var frm = document.frm;
	 
		var startday = frm.ChkStart.value;
		var endday = frm.ChkEnd.value;
		var usePersonNo = frm.usePersonNo.value; 
		document.getElementById("ifmProc").src="procCalPoint.asp?ChkStart="+startday+"&ChkEnd="+endday+"&usePersonNo="+usePersonNo+"&empno=<%=empno%>";
	}
	
	function jschkPoint(){
		document.frm.chkp.value =0 ;
	}
	
 
//-->
</script>
 
<form name="frm" method="post" action="tenbyten_agit_Process.asp" >
<input type="hidden" name="mode" value="<%=chkIIF(idx="","add","modi")%>">
<input type="hidden" name="idx" value="<%=idx%>">
<table width="100%" border="0" cellpadding="5" cellspacing="0" class="a">
<tr>
	<td><b>�ٹ����� ����Ʈ ����</b><br><hr width="100%"></td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="5" cellspacing="1" class="a" bgcolor="#909090">
			<tr bgcolor="#FFFFFF">
			<td width="120" bgcolor="<%=adminColor("sky")%>" align="center"><b>���</b></td>
			<td><%=empno%>
				<input type="hidden" name="sEn" size="20" class="text" value="<%=empno%>"> 
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td width="120" bgcolor="<%=adminColor("sky")%>" align="center"><b>����ID</b></td>
			<td> <%=userid%>
				<input type="hidden" name="userid" size="20" class="text" value="<%=userid%>">
				<!--<input type="button" value="IDȮ��" class="button_s" style="width:55px;text-align:center;" onclick="chkTenMember(frm.userid.value)">
				<font color=gray>(ID �Ǵ� �̸� �Է�)</font>
				<input type="hidden" name="chkCfm" value="<%=chkIIF(idx="","N","Y")%>">-->
			</td>
		</tr>
		<% if C_ADMIN_AUTH or C_PSMngPart then %>
		<tr bgcolor="#FFFFFF">
			<td width="120" bgcolor="<%=adminColor("sky")%>" align="center"><b>�̸�/����</b></td>
			<td> <%=susername%>/<%=sposit_name%>  <input type="hidden" name="username" value="<%=susername%>"></td>
		</tr>
		<% else %>
		<tr bgcolor="#FFFFFF">
			<td width="120" bgcolor="<%=adminColor("sky")%>" align="center"><b>�̸�</b></td>
			<td> <%=susername%> <input type="hidden" name="username" value="<%=susername%>"></td>
		</tr>
		<% end if %>
		<tr bgcolor="#FFFFFF">
			<td width="120" bgcolor="<%=adminColor("sky")%>" align="center"><b>�ҼӺμ�</b></td>
			<td><%=sdepartmentNameFull %></td>
		</tr> 
		<tr bgcolor="#FFFFFF">
			<td width="120" bgcolor="<%=adminColor("sky")%>" align="center"><b>��󿬶�ó</b></td>
			<td>
				��ȭ <input type="text" name="userPhone" size="18" maxlength="18" class="text" value="<%=suserPhone%>"> /
				�޴��� <input type="text" name="userHP" size="18" maxlength="18" class="text" value="<%=suserCell%>">
			</td>
		</tr>	
		<tr bgcolor="#FFFFFF">
			<td width="120" bgcolor="<%=adminColor("sky")%>" align="center"><b>����Ʈ</b></td>
			<td>
				<select name="AreaDiv">
					<option value="">::����::</option>
					<!--<option value="1">���ֵ�</option>-->
					<!--<option value="2">����</option>-->
					<option value="3">����</option>
				</select>
				<script type="text/javascript">
					document.frm.AreaDiv.value="<%=AreaDiv%>";
				</script>
			</td>
		</tr>
		
		<tr bgcolor="#FFFFFF">
			<td width="120" bgcolor="<%=adminColor("sky")%>" align="center"><b>�̿�Ⱓ</b></td>
			<td style="line-height:18px;">
				<input id="ChkStart" name="ChkStart" value="<%=sDt%>" class="text" onchange="chkTerm()" size="10" maxlength="10" <%if idx <> "" then%>readonly<%end if%>/><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="ChkStart_trigger" border="0" style="cursor:pointer" align="absmiddle" />
				<input type="text" name="ChkSTime" size="5" maxlength="5" class="text" value="<%=sTm%>" readonly > ~
				<input id="ChkEnd" name="ChkEnd" value="<%=eDt%>" class="text" onchange="chkTerm()" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="ChkEnd_trigger" border="0" style="cursor:pointer" align="absmiddle" />
				<input type="text" name="ChkETime" size="5" maxlength="5" class="text" value="<%=eTm%>" readonly >
		    	<font color=gray>(<input type="text" name="uTerm" readonly class="text" value="<%=DateDiff("d",sDt,eDt)%>" style="text-align:right; width:20px; border:0px; color:gray;">��)</font>
				<script language="javascript">
					var CAL_Start = new Calendar({
						inputField : "ChkStart", trigger    : "ChkStart_trigger",
						onSelect: function() {
							var date = Calendar.intToDate(this.selection.get());
							CAL_End.args.min = date;
							CAL_End.redraw();
							chkTerm();
							this.hide();
						}, bottomBar: true, dateFormat: "%Y-%m-%d"
					});
					var CAL_End = new Calendar({
						inputField : "ChkEnd", trigger    : "ChkEnd_trigger",
						onSelect: function() {
							var date = Calendar.intToDate(this.selection.get());
							CAL_Start.args.max = date;
							CAL_Start.redraw();
							chkTerm();
							this.hide();
						}, bottomBar: true, dateFormat: "%Y-%m-%d"
					});
				</script>
			</td>
		</tr>
		<%dim iPoint, sMoney,avPoint%>
		<tr bgcolor="#FFFFFF">
			<td width="120" bgcolor="<%=adminColor("sky")%>" align="center"><b>�����ο�</b></td>
			<td><input type="text" name="usePersonNo" size="4" class="text" value="<%=usePersonNo%>" onKeyPress="jschkPoint();" style="text-align:right;padding-right:3px;">�� <font color=gray>(�������� ���ο���)</font>
				
				</td>
		</tr> 
		<%if idx ="" then%>
		<tr>
			<td colspan="2" bgcolor="#FFFFFF" align="center"><input type="button" class="button" value="����Ʈ/ �ݾ� Ȯ��" onClick="jsSetPoint();">
				<input type="hidden" name="chkp" id="chkp" value="0"></td>
		</tr>
		<%end if%>
		<tr bgcolor="#FFFFFF">
			<td width="120" bgcolor="<%=adminColor("sky")%>" align="center"><b>����Ʈ</b></td>
			<td> 
				<input type="text" name="iPoint" id="iPoint" size="4" class="text_ro" value="<%=usepoint%>" style="text-align:right;padding-right:3px;" readonly>
				<%if idx ="" then%>
				/ 
				<input type="text" name="avPoint" id="avPoint" value="<%=totap-useap%>" readonly style="border:0px;width:30px;" class="text">
				<font color=gray>(��뿹������Ʈ/��밡������Ʈ)</font>
				<%end if%>
				<div style="color:blue;font-size:11px;padding:3px">���� 1����Ʈ, ��,��, ������ ���� , �������� 2����Ʈ ����</div>
				</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td width="120" bgcolor="<%=adminColor("sky")%>" align="center"><b>�ݾ�</b></td>
			<td><input type="text" name="sMoney" id="sMoney" size="10" class="text_ro" value="<%=useMoney%>" style="text-align:right;padding-right:3px;" readonly>��
				<span style="color:blue;font-size:11px;">1��:15,000��/ 5���̻� 1���� �߰�</span>
				</td>
			
		</tr>
		<tr bgcolor="#FFFFFF">
			<td width="120" bgcolor="<%=adminColor("sky")%>" align="center"><b>���</b></td>
			<td><textarea name="etcComment" class="textarea" style="width:100%; height:50px;"><%=etcComment%></textarea></td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td colspan="2" align="center">		 
	<% if idx<>"" then
					if  empno = session("ssBctSn") and sDt >= date() then %>
				<input type="button" value="��û���" class="button" style="width:60px;text-align:center;" onclick="delBook()">
 <% 			end if
		else %>
				<input type="botton" value="��û" class="button" style="width:60px;text-align:center;" onClick="jsSubmit();">
	<% end if %>
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
</form>
<iframe id="ifmProc" src="" width="0" height="0" frameborder="0"></iframe>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->