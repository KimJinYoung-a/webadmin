<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �������� ��༭
' History : 2011.01.12 ������  ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenMemberCls.asp" -->
<!-- #include virtual="/lib/classes/admin/tenbyten/TenByTenPayCls.asp" -->
<!-- #include virtual="/admin/eventmanage/common/event_function_v3.asp"-->
<%
Dim cMember,clsPayForm
Dim sEmpno , ino
Dim susername, susermail, sdirect070, djoinday, blnstatediv, spart_name, sposit_name, sjob_name
Dim startdate, enddate,defaultpay ,foodpay,jobpay ,inBreakTime  , holidaywdtime	,regdate    ,lastupdate ,adminid,iposit_sn,dretireday,sjuminno,suserphone,susercell,szipcode,szipaddr,suseraddr
Dim StartHour(8), StartMinute(8), EndHour(8), EndMinute(8), BreakSHour(8), BreakSMinute(8),  BreakEHour(8), BreakEMinute(8),DutyTime(8) ,NightTime(8), iworktype(8)
Dim totDutyTime,iOverTime,iPatternSeq,part_sn,spatternname,totNightTime, iHolidayTime,avgWeek,totPaySum
Dim iTotCnt,iPageSize, iTotalPage,page
Dim arrList, intLoop, jobkind, placekind

avgWeek = 4.345238095
sEmpno =   requestCheckvar(request("sEN"),14)
ino =requestCheckvar(request("ino"),10)

'��� ������� ��������-----------------
Set cMember  = new CTenByTenMember
	cMember.Fempno		= sEmpno
	cMember.fnGetMemberData
	susername	= cMember.Fusername
	sjuminno		= cMember.Fjuminno
	suserphone	= cMember.FuserPhone
	susercell		= cMember.Fusercell
	szipcode		= cMember.Fzipcode
	szipaddr		= cMember.Fzipaddr
	suseraddr	= cMember.Fuseraddr
	djoinday	  	= cMember.Fjoinday
	blnstatediv 	= cMember.Fstatediv
	iposit_sn		= cMember.Fposit_sn
	spart_name  	= cMember.Fpart_name
	sposit_name 	= cMember.Fposit_name
	sjob_name	= cMember.Fjob_name
	dretireday		= cMember.Fretireday
Set cMember = nothing
'---------------------------------------
'��� ������� ��������-----------------
Set clsPayForm = new CPayForm
	clsPayForm.Fempno= sEmpno
	clsPayForm.Fino = ino
	clsPayForm.fnGetDefaultPayData

	startdate		= clsPayForm.Fstartdate
	enddate		= clsPayForm.Fenddate

	defaultpay    	= clsPayForm.Fdefaultpay
	foodpay	    	= clsPayForm.Ffoodpay
	jobpay		= clsPayForm.Fjobpay

	inBreakTime	= clsPayForm.FinBreakTime
	iOverTime		= clsPayForm.FOverTime

	For intLoop = 1 To 7
	StartHour(intLoop) 		= clsPayForm.FStartHour(intLoop)
	StartMinute(intLoop)  	= clsPayForm.FStartMinute(intLoop)
	EndHour(intLoop)       	= clsPayForm.FEndHour(intLoop)
	EndMinute(intLoop)       = clsPayForm.FEndMinute(intLoop)
	BreakSHour(intLoop)     	= clsPayForm.FBreakSHour(intLoop)
	BreakSMinute(intLoop)     = clsPayForm.FBreakSMinute(intLoop)
	BreakEHour(intLoop)     	= clsPayForm.FBreakEHour(intLoop)
	BreakEMinute(intLoop)     = clsPayForm.FBreakEMinute(intLoop)
	DutyTime(intLoop)		=  clsPayForm.FDutyTime(intLoop)
	iworktype(intLoop)		= clsPayForm.Fworktype(intLoop)
	Next

	totDutyTime  = clsPayForm.FTotDutyTime
	totNightTime	= clsPayForm.FtotNightTime
	totPaySum	=clsPayForm.FTotPaySum

	holidaywdtime	  = clsPayForm.Fholidaywdtime
	regdate        =clsPayForm.Fregdate
	lastupdate     =clsPayForm.Flastupdate
	adminid        =clsPayForm.Fadminid
	jobkind		= clsPayForm.Fjobkind
	placekind		= clsPayForm.Fplacekind
Set clsPayForm = nothing
'---------------------------------------
%>
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link rel="stylesheet" href="/css/scm.css" type="text/css">
<script language="javascript">
<!--
	 document.body.onload=function(){window.print();}
//-->
</script>
</head>
<body leftmargin="0" topmargin="0">
<table width="100%" border="0" cellpadding="3" cellspacing="0" class="a">
<tr>
	<td align="center" style="font-family: �������,Verdana;font-size:18px;"><strong>�� �� �� �� �� �� �� �� [ <%IF iposit_sn =13 THEN%>�� ��<%ELSE%>�� ��<%END IF%> ]</strong></td>
</tr>
<tr>
	<td>
		<table width="100%" border="1" cellpadding="3" cellspacing="0" align="center" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr align="center">
			<td rowspan="2" bgcolor="#FFFFFF" valign="top"><b>�����<br>(��)</b></td>
			<td bgcolor="<%= adminColor("tabletop") %>"><b>��ȣ</b></td>
			<td bgcolor="#FFFFFF">���ٹ�����</td>
			<td bgcolor="<%= adminColor("tabletop") %>"><b>����ڹ�ȣ</b></td>
			<td bgcolor="#FFFFFF">211-87-00620</td>
			<td bgcolor="<%= adminColor("tabletop") %>"><b>��ǥ��</b></td>
			<td bgcolor="#FFFFFF">������</td>
		</tr>
		<tr  align="center">
			<td bgcolor="<%= adminColor("tabletop") %>"><b>�ּ�</b></td>
			<td colspan="6" bgcolor="#FFFFFF">(03082) ����� ���α� ���з� 57 ȫ�ʹ��б� ���з�ķ�۽� ������ 14�� �ٹ�����</td>
		</tr>
		<tr align="center">
			<td rowspan="3" bgcolor="#FFFFFF" valign="top"><b>�ٷ���<br>(��)</b></td>
			<td bgcolor="<%= adminColor("tabletop") %>"><b>����</b></td>
			<td bgcolor="#FFFFFF"><%=susername%></td>
			<td bgcolor="<%= adminColor("tabletop") %>"><b>�ֹε�Ϲ�ȣ</b></td>
			<td bgcolor="#FFFFFF"><%=LEFT(sjuminno,8)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td><!-- eastone ����-->
			<td bgcolor="<%= adminColor("tabletop") %>"><b>��ȭ��ȣ</b></td>
			<td bgcolor="#FFFFFF"><%IF susercell <> "" THEN%><%=susercell%><%ELSE%><%=suserphone%><%END IF%></td>
		</tr>
		<tr  align="center">
			<td bgcolor="<%= adminColor("tabletop") %>"><b>�ּ�</b></td>
			<td colspan="6" bgcolor="#FFFFFF"><%=szipaddr%> <%=suseraddr%></td>
		</tr>
		<tr  align="center">
			<td bgcolor="<%= adminColor("tabletop") %>"><b>����</b></td>
			<td colspan="2" bgcolor="#FFFFFF"><% GetEvnetKindName "jobkind", jobkind %></td>
			<td bgcolor="<%= adminColor("tabletop") %>"><b>�ٹ���</b></td>
			<td colspan="2" bgcolor="#FFFFFF"><% GetEvnetKindName "placekind", placekind %></td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td  align="center">���� ���ٹ�����(���� "��")�� �ٷ���(���� "��")(��)�� ��ȣ ������ �������� �����ǻ翡 ����<br />
		������ ���� �ٷΰ���� ü���ϰ� ��ȣ ������ ���� �� �ؼ��� ���� �����մϴ�.
	</td>
</tr>
<tr>
	<td>
		<table width="100%" border="0" cellpadding="2" cellspacing="0" align="center" class="a">
		<tr>
			<td valign="top" width="85"><b>1.���Ⱓ</b></td>
			<td>
				<table width="100%" border="0" cellpadding="0" cellspacing="0" align="center" class="a">
				<tr>
					<td>��������:</td>
					<td><b><%=year(startdate)%>�� <%=month(startdate)%>�� <%=day(startdate)%>��</b></td>
					<td>���������:</td>
					<td><b><%=year(enddate)%>�� <%=month(enddate)%>�� <%=day(enddate)%>��</b></td>
				</tr>
				<tr>
					<td colspan="4">
						�ٷΰ�� ���� 1���� �� �� ����� ���� Ư���� �ǻ簡ǥ�ð� ���� �� �ڵ������ȴ�.<br />
						�ӱ� ���� ���� �߻� ��(�����ӱ� ���, �ٷ����� ���� ��), ���� ����� �����Ѵ�.
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td><b>2.�ÿ�Ⱓ</b></td>
			<td>�ű� �Ի����� ��� ���� �Ի��Ϸκ��� 3�������� �ÿ�Ⱓ���� �Ѵ�.</td>
		</tr>
		<tr>
			<td><b>3.����/�ٹ���</b></td>
			<td>���� �λ�߷ɿ� ���� �����ϴ� ������ �ٹ����� �Ѵ�. </td>
		</tr>
		<tr>
			<td valign="top"><b>4.�ٷ�����</b></td>
			<td>
				<table width="100%" border="0" cellpadding="3" cellspacing="0" align="center" class="a">
				<tr>
					<td  valign="top" width="70"><b>1) �޿�����</b></td>
					<td width="500">
				<%IF iposit_sn =13 THEN%>
					�ñ޾�( <b><%=formatnumber(defaultpay,0)%></b>��) X �ٹ��ð� �� ������ �����ϸ�, ���޼����� �߻��ϴ� ���,<br>
					�̸� �߰��Ͽ� �����Ѵ�.<br>
					��, �ٷα��ع��� �ǰ� 4�ָ� ����� 1�� �����ٷνð��� 15�ð� �̸��� ��쿡��<br>
					�������� �ο����� �ƴ��Ѵ�.
				<%END IF%>
					</td>
				</tr>
				<%IF iposit_sn=12 or iposit_sn=15 THEN%>
				<tr>
					<td  colspan="2">
						<table width="100%" border="1" cellpadding="3" cellspacing="0" align="center" class="a" bgcolor=#BABABA>
						<tr align="center">
							<td  bgcolor="<%= adminColor("tabletop") %>" width="15%">�⺻��</td>
							<td  bgcolor="<%= adminColor("tabletop") %>" width="15%">���޼���</td>
							<td  bgcolor="<%= adminColor("tabletop") %>" width="15%">�ð��ܼ���</td>
							<td  bgcolor="<%= adminColor("tabletop") %>" width="15%">�߰��ٹ�����</td>
							<td  bgcolor="<%= adminColor("tabletop") %>" width="15%">&nbsp;</td>
							<td  bgcolor="<%= adminColor("tabletop") %>" width="25%">���޿�</td>
						</tr>
						<tr  align="center">
							<td  bgcolor="#FFFFFF"><%IF totDutyTime =0 THEN%>&nbsp;<%ELSE%><%=formatnumber(defaultpay*ceilValue(totDutyTime/60*avgWeek),0)%><%END IF%></td>
							<td  bgcolor="#FFFFFF"><%IF holidaywdtime =0 THEN%>&nbsp;<%ELSE%><%=formatnumber(defaultpay*ceilValue(holidaywdtime/60*avgWeek),0)%><%END IF%></td>
							<td  bgcolor="#FFFFFF"><%IF iOverTime =0 THEN%>&nbsp;<%ELSE%><%=formatnumber(defaultpay*iOverTime*1.5,0)%><%END IF%></td>
							<td  bgcolor="#FFFFFF"><%IF totNightTime =0 THEN%>&nbsp;<%ELSE%><%=formatnumber(defaultpay*ceilValue(totNightTime/60*avgWeek)*0.5,0)%><%END IF%></td>
							<td  bgcolor="#FFFFFF">&nbsp;</td>
							<td  bgcolor="#FFFFFF"><%IF totPaySum =0 THEN%>&nbsp;<%ELSE%><%=formatnumber(totPaySum,0)%><%END IF%></td>
						</tr>
						</table>
						<% if sEmpno ="90201704030065"or sEmpno="90201702010023" or sEmpno= "90201602150020" or sEmpno ="90201602150021" or sEmpno = "90201702010024" or sEmpno = "90201702010025" or sEmpno = "90201704120084" or sEmpno = "90201702010026" then '//���⿹��ó�� 2017.12 ���������� ��û%>
						<p style="padding:1px">�� ����� �� 22�ð��� ����ٷμ����� �ջ�� ���������ӱ��� �ƴ��� Ȯ���Ͽ����� ���� �̿� �����Ѵ�.�� ����ٷ� � ���ؼ��� ������ ��������� �����Ѵ�.</p>
					 <%end if%> 
					</td>
				</tr>
				<%END IF%>
				<tr>
					<td  colspan="2"><b>2) �ٷνð�</b><br>
						<table width="100%" border="1" cellpadding="3" cellspacing="0" align="center" class="a" bgcolor=#BABABA>
						<tr align="center">
							<td  bgcolor="<%= adminColor("tabletop") %>" rowspan="2">����</td>
							<td  bgcolor="<%= adminColor("tabletop") %>" colspan="2">�ٹ��ð�</td>
							<td  bgcolor="<%= adminColor("tabletop") %>" colspan="2">�ްԽð�</td>
							<td  bgcolor="<%= adminColor("tabletop") %>" rowspan="2">���</td>
						</tr>
						<tr align="center">
							<td  bgcolor="<%= adminColor("tabletop") %>" >����</td>
							<td  bgcolor="<%= adminColor("tabletop") %>" >����</td>
							<td  bgcolor="<%= adminColor("tabletop") %>" >����</td>
							<td  bgcolor="<%= adminColor("tabletop") %>" >����</td>
						</tr>
						<%
						For intLoop = 1 To 7%>
						<tr align="center" bgcolor="#FFFFFF">
							<td><%=fnGetStringWD(intLoop)%></td>
							<td><%IF StartHour(intLoop) ="00" and StartMinute(intLoop) ="00" THEN%>&nbsp;<%ELSE%><%=StartHour(intLoop)%>:<%=StartMinute(intLoop)%><%END IF%></td>
							<td><%IF EndHour(intLoop) ="00" and EndMinute(intLoop) ="00" THEN%>&nbsp;<%ELSE%><%=EndHour(intLoop)%>:<%=EndMinute(intLoop)%><%END IF%></td>
							<td><%IF BreakSHour(intLoop) ="00" and BreakSMinute(intLoop) ="00" THEN%>&nbsp;<%ELSE%><%=BreakSHour(intLoop)%>:<%=BreakSMinute(intLoop)%><%END IF%></td>
							<td><%IF BreakEHour(intLoop) ="00" and BreakEMinute(intLoop) ="00" THEN%>&nbsp;<%ELSE%><%=BreakEHour(intLoop)%> : <%=BreakEMinute(intLoop)%><%END IF%> </td>
							<td><%IF iworktype(intLoop)= "1"  THEN%>
								�ٹ���
								<%ELSEIF  iworktype(intLoop)= "2"  THEN%>
									��������
								<%ELSEIF  iworktype(intLoop)= "3"  THEN%>
									������
								<%ELSEIF iworktype(intLoop)  = "4" THEN%>
						 		��������
								<%END IF%>
							</td>
						</tr>
						<%  	Next %>
						</table>
					<td>
				</tr>
				<tr>
					<td valign="top"><b>3) ���޽ñ�</b></td>
					<td><%IF iposit_sn =12 THEN%>
						�ſ� 1�Ϻ��� ���ϱ����� �ӱ��� ��� ����
						<%ELSE%>
						������� �ӱ��� �Ϳ� 5��
						<%END IF%>
						�� �����ϸ�,<br>
						������������ ��õ¡�� �� �� �����Ѵ�.
					</td>
				</tr>
				<tr>
					<td valign="top"><b>4) �����޿�</b></td>
					<td>�ٷα��ع��� ���� ��ӱټӳ�� 1�⿡ ���Ͽ� 30�Ϻ��� ����ӱ����� �����Ѵ�.<br>
						���� ������ �����Ͽ� �߻��ϴ� ��� ��ǰ�� �������� ���� ���� �Ϳ� 15���� ���޹޴� �Ϳ� �����Ѵ�.
					</td>
				</tr>
				<tr>
					<td valign="top"><b>5) ��������</b></td>
					<td>���ܰ�� 3ȸ �̻� �߻��� ȸ��� �ذ���ġ �� �� �ִ�.��Ÿ �ٸ� ������ �����Ģ �� ��Կ� �ǰ��Ѵ�.</td>
				</tr>
				<tr>
					<td valign="top"><b>6) ��������</b></td>
					<td>�ٷ����� ��(5/1), ������(�߻���)�� �ϰ�, �����ް��� 1���̻� �ٹ��� 15����<br>
						�ο��ϸ�, �� �� ������ �ٷα��ع��� ������.
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td valign="top"><b>5. ��Ÿ����</b></td>
			<td>�� ��༭�� ��õ��� ���� ���׿� ���ؼ��� �ٷα��ع� ���� ������� �Ǵ� �����Ģ,��Ե� ����<br>
			  ������ ���� ������ ������.
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td align="center">���� ���� �ٷΰ���� ü����.</td>
</tr>
<tr>
	<td align="center"> <%=year(startdate)%> �� &nbsp;&nbsp; <%=month(startdate)%> �� &nbsp;&nbsp; <%=day(startdate)%> ��</td>
</tr>
<tr>
	<td align="center"  >
		<table width="100%" border="1" cellpadding="3" cellspacing="0" align="center" class="a" bgcolor="<%= adminColor("tablebg") %>">
		<tr bgcolor="#FFFFFF" align="center">
			<td width="10%"><b>�����<br>(��)</b></td>
			<td width="40%">(��)�ٹ����� ��ǥ�̻� ������ <img src="http://scm.10x10.co.kr/images/seal1.gif" width="80" align="absmiddle"></td>
			<td width="10%"><b>�ٷ���<br>(��)</b></td>
			<td align="right">(��)&nbsp;&nbsp;&nbsp;&nbsp;</td>
		</tr>
		</table>
	</td>
</tr>
</html>
</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->