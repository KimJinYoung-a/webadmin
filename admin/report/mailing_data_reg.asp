<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  ���� ���
' History : 2007.08.27 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/report/reportcls.asp"-->
<%
Dim omd
Dim idx,mode
	idx = requestcheckvar(getNumeric(request("idx")),10)
	mode = requestcheckvar(request("mode"),32)

If idx = "" Then idx=0

set omd = New CMailzineOne
	omd.GetMailingOne idx
%>
<link href="/css/report.css" rel="stylesheet" type="text/css">
<script type="text/javascript">

function TnMailDataReg(frm){
	if(frm.title.value == ""){
		alert("�߼��̸��� �����ּ���");
		frm.title.focus();
	}
	else if(frm.gubun.value == ""){
		alert("�߼۱����� �����ּ���");
		frm.gubun.focus();
	}
	else if(frm.startdate.value == ""){
		alert("�߼۽��۽ð��� �����ּ���");
		frm.startdate.focus();
	}
	else if(frm.enddate.value == ""){
		alert("�߼�����ð��� �����ּ���");
		frm.enddate.focus();
	}
	else if(frm.reenddate.value == ""){
		alert("��߼�����ð��� �����ּ���");
		frm.reenddate.focus();
	}
	else if(frm.totalcnt.value == ""){
		alert("�Ѵ���ڼ��� �����ּ���");
		frm.totalcnt.focus();
	}
	else if(frm.realcnt.value == ""){
		alert("�ǹ߼������ �����ּ���");
		frm.realcnt.focus();
	}
	else if(frm.realpct.value == ""){
		alert("�ǹ߼ۺ����� �����ּ���");
		frm.realpct.focus();
	}
	else if(frm.filteringcnt.value == ""){
		alert("���͸������ �����ּ���");
		frm.filteringcnt.focus();
	}
	else if(frm.filteringpct.value == ""){
		alert("���͸������� �����ּ���");
		frm.filteringpct.focus();
	}
	else if(frm.successcnt.value == ""){
		alert("�����߼������ �����ּ���");
		frm.successcnt.focus();
	}
	else if(frm.successpct.value == ""){
		alert("�������� �����ּ���");
		frm.successpct.focus();
	}
	else if(frm.failcnt.value == ""){
		alert("���й߼������ �����ּ���");
		frm.failcnt.focus();
	}
	else if(frm.failpct.value == ""){
		alert("�������� �����ּ���");
		frm.failpct.focus();
	}
	else if(frm.opencnt.value == ""){
		alert("��������� �����ּ���");
		frm.opencnt.focus();
	}
	else if(frm.openpct.value == ""){
		alert("�������� �����ּ���");
		frm.openpct.focus();
	}
	else if(frm.noopencnt.value == ""){
		alert("�̿�������� �����ּ���");
		frm.noopencnt.focus();
	}
	else if(frm.noopenpct.value == ""){
		alert("�̿������� �����ּ���");
		frm.noopenpct.focus();
	}
	else{
		frm.submit();
	}
}

</script>

<form method="post" name="sform" action="/admin/report/domaildata.asp" style="margin:0px;">
<input type="hidden" name="mode" value="<% = mode %>">
<input type="hidden" name="idx" value="<% = idx %>">
<table cellpadding="0" cellspacing="0" border="0">
<tr>
	<td>
		<table width="660" border="0" cellspacing="0" cellpadding="0">
		<tr>
			<td height="2" bgcolor="C6E7EA"></td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<table width="1" border="0" cellspacing="0" cellpadding="0">
		<tr>
			<td align="center" bgcolor="#DDDDDD" class="PD1px">
				<table width="658" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td bgcolor="#ffffff">
						<table width="658" border="0" bordercolordark="white"  cellspacing="0" cellpadding="0">
						<tr>
							<td width="170" height="30" align="center" bgcolor="EFF5F1"><font color="57645B"><strong>�߼۱���</strong></font></td>
							<td width="20">&nbsp;</td>
							<td width="478" ><input name="gubun" size="35" class="box" type="text" value="<% = omd.fgubun %>"> ex) mailzine , mailzine_not , mailzine_event</td>
						</tr>
						</table>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td align="center" bgcolor="#DDDDDD" class="PD1px">
				<table width="658" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td bgcolor="#ffffff">
						<table width="658" border="0" bordercolordark="white"  cellspacing="0" cellpadding="0">
						<tr>
							<td width="170" height="30" align="center" bgcolor="EFF5F1"><font color="57645B"><strong>�߼��̸�</strong></font></td>
							<td width="20">&nbsp;</td>
							<td width="478" ><input name="title" size="65" class="box" type="text" value="<% = omd.Ftitle %>"></td>
						</tr>
						</table>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td align="center" bgcolor="#DDDDDD">
				<table width="658" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td bgcolor="#ffffff">
						<table width="658" border="0" bordercolordark="white"  cellspacing="0" cellpadding="0">
						<tr>
							<td width="170" height="30" align="center" bgcolor="EFF5F1"><font color="57645B"><strong>�߼۽��۽ð�</strong></font></td>
							<td width="20">&nbsp;</td>
							<td width="478" ><input name="startdate" size="65" class="box" type="text" value="<% = omd.Fstartdate %>"></td>
						</tr>
						</table>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td align="center" bgcolor="#DDDDDD" class="PD1px">
				<table width="658" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td bgcolor="#ffffff">
						<table width="658" border="0" bordercolordark="white"  cellspacing="0" cellpadding="0">
						<tr>
							<td width="170" height="30" align="center" bgcolor="EFF5F1"><font color="57645B"><strong>�߼�����ð�</strong></font></td>
							<td width="20">&nbsp;</td>
							<td width="478" ><input name="enddate" size="65" class="box" type="text" value="<% = omd.Fenddate %>"></td>
						</tr>
						</table>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td align="center" bgcolor="#DDDDDD">
				<table width="658" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td bgcolor="#ffffff">
						<table width="658" border="0" bordercolordark="white"  cellspacing="0" cellpadding="0">
						<tr>
							<td width="170" height="30" align="center" bgcolor="EFF5F1"><font color="57645B"><strong>��߼�����ð�</strong></font></td>
							<td width="20">&nbsp;</td>
							<td width="478" ><input name="reenddate" size="65" class="box" type="text" value="<% = omd.Freenddate %>"></td>
						</tr>
						</table>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td align="center" bgcolor="#DDDDDD" class="PD1px">
				<table width="658" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td bgcolor="#ffffff">
						<table width="658" border="0" bordercolordark="white"  cellspacing="0" cellpadding="0">
						<tr>
							<td width="170" height="30" align="center" bgcolor="EFF5F1"><font color="57645B"><strong>�Ѵ���ڼ�</strong></font></td>
							<td width="20">&nbsp;</td>
							<td width="478" ><input name="totalcnt" size="30" class="box" type="text" value="<% = omd.Ftotalcnt %>"></td>
						</tr>
						</table>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td align="center" bgcolor="#DDDDDD">
				<table width="658" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td bgcolor="#ffffff">
						<table width="658" border="0" bordercolordark="white"  cellspacing="0" cellpadding="0">
						<tr>
							<td width="170" height="30" align="center" bgcolor="EFF5F1"><font color="57645B"><strong>�ǹ߼����(�߼ۺ���)</strong></font></td>
							<td width="20">&nbsp;</td>
							<td width="478" ><input name="realcnt" size="20" class="box" type="text" value="<% = omd.Frealcnt %>">&nbsp;&nbsp;<input name="realpct" size="20" class="box" type="text" value="<% = omd.Frealpct %>"></td>
						</tr>
						</table>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td align="center" bgcolor="#DDDDDD" class="PD1px">
				<table width="658" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td bgcolor="#ffffff">
						<table width="658" border="0" bordercolordark="white"  cellspacing="0" cellpadding="0">
						<tr>
							<td width="170" height="30" align="center" bgcolor="EFF5F1"><font color="57645B"><strong>���͸� ���(���͸� ����)</strong></font></td>
							<td width="20">&nbsp;</td>
							<td width="478" ><input name="filteringcnt" size="20" class="box" type="text" value="<% = omd.Ffilteringcnt %>">&nbsp;&nbsp;<input name="filteringpct" size="20" class="box" type="text" value="<% = omd.Ffilteringpct %>"></td>
						</tr>
						</table>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td align="center" bgcolor="#DDDDDD">
				<table width="658" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td bgcolor="#ffffff">
						<table width="658" border="0" bordercolordark="white"  cellspacing="0" cellpadding="0">
						<tr>
							<td width="170" height="30" align="center" bgcolor="EFF5F1"><font color="57645B"><strong>�����߼� ���(������)</strong></font></td>
							<td width="20">&nbsp;</td>
							<td width="478" ><input name="successcnt" size="20" class="box" type="text" value="<% = omd.Fsuccesscnt %>">&nbsp;&nbsp;<input name="successpct" size="20" class="box" type="text" value="<% = omd.Fsuccesspct %>"></td>
						</tr>
						</table>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td align="center" bgcolor="#DDDDDD" class="PD1px">
				<table width="658" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td bgcolor="#ffffff">
						<table width="658" border="0" bordercolordark="white"  cellspacing="0" cellpadding="0">
						<tr>
							<td width="170" height="30" align="center" bgcolor="EFF5F1"><font color="57645B"><strong>���й߼� ���(������)</strong></font></td>
							<td width="20">&nbsp;</td>
							<td width="478" ><input name="failcnt" size="20" class="box" type="text" value="<% = omd.Ffailcnt %>">&nbsp;&nbsp;<input name="failpct" size="20" class="box" type="text" value="<% = omd.Ffailpct %>"></td>
						</tr>
						</table>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td align="center" bgcolor="#DDDDDD">
				<table width="658" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td bgcolor="#ffffff">
						<table width="658" border="0" bordercolordark="white"  cellspacing="0" cellpadding="0">
						<tr>
							<td width="170" height="30" align="center" bgcolor="EFF5F1"><font color="57645B"><strong>���� ���(������)</strong></font></td>
							<td width="20">&nbsp;</td>
							<td width="478" ><input name="opencnt" size="20" class="box" type="text" value="<% = omd.Fopencnt %>">&nbsp;&nbsp;<input name="openpct" size="20" class="box" type="text" value="<% = omd.Fopenpct %>"></td>
						</tr>
						</table>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td align="center" bgcolor="#DDDDDD" class="PD1px">
				<table width="658" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td bgcolor="#ffffff">
						<table width="658" border="0" bordercolordark="white"  cellspacing="0" cellpadding="0">
						<tr>
							<td width="170" height="30" align="center" bgcolor="EFF5F1"><font color="57645B"><strong>�̿��� ���(�̿�����)</strong></font></td>
							<td width="20">&nbsp;</td>
							<td width="478" ><input name="noopencnt" size="20" class="box" type="text" value="<% = omd.Fnoopencnt %>">&nbsp;&nbsp;<input name="noopenpct" size="20" class="box" type="text" value="<% = omd.Fnoopenpct %>"></td>
						</tr>
						</table>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<!--2016-12-07 ���¿� �߰�-->
		<tr>
			<td align="center" bgcolor="#DDDDDD">
				<table width="658" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td bgcolor="#ffffff">
						<table width="658" border="0" bordercolordark="white"  cellspacing="0" cellpadding="0">
						<tr>
							<td width="170" height="30" align="center" bgcolor="EFF5F1"><font color="57645B"><strong>Ŭ�� ��(Ŭ����)</strong></font></td>
							<td width="20">&nbsp;</td>
							<td width="478" ><input name="clickcnt" size="20" class="box" type="text" value="<% = omd.Fclickcnt %>">&nbsp;&nbsp;<input name="clickpct" size="20" class="box" type="text" value="<% = omd.Fclickpct %>"></td>
						</tr>
						</table>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<!-- //-2016-12-07 ���¿� �߰�-->
		<tr>
			<td align="center" bgcolor="#DDDDDD">
				<table width="658" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td bgcolor="#ffffff">
						<table width="658" border="0" bordercolordark="white"  cellspacing="0" cellpadding="0">
						<tr>
							<td width="170" height="30" align="center" bgcolor="EFF5F1"><font color="57645B"><strong>���Ϸ�</strong></font></td>
							<td width="20">&nbsp;</td>
							<td width="478" >
								<%= omd.fmailergubun %>
								<Br>KEY : <%= omd.fmailer_key_maeching %>
							</td>
						</tr>
						</table>
					</td>
				</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td align="center" bgcolor="#DDDDDD" style="padding-bottom:1px">
				<table width="658" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td bgcolor="#ffffff" align="right"><a href="javascript:TnMailDataReg(sform);">�����ϱ�</a>&nbsp;&nbsp;&nbsp;</td>
				</tr>
				</table>
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
</form>

<script type="text/javascript">

function autochk(){
	var arrFrm = new Array();  //ã�� ����

	arrFrm[0] = new Array() //���� ���� ���� �Է�
	arrFrm[1] = new Array()	//�ʵ�

	arrFrm[0][0]	=	'�߼��̸�';
	arrFrm[0][1]	=	'�߼۽��۽ð�';
	arrFrm[0][2]	=	'�߼�����ð�';
	arrFrm[0][3]	=	'��߼�����ð�';
	arrFrm[0][4]	=	'�Ѵ���ڼ�';
	arrFrm[0][5]	=	'�ǹ߼����';
	arrFrm[0][6]	=	'���͸����';
	arrFrm[0][7]	=	'�����߼����';
	arrFrm[0][8]	=	'���й߼����';
	arrFrm[0][9]	=	'�������';
	arrFrm[0][10]	=	'�̿������';

	arrFrm[1][0]	=	'title';
	arrFrm[1][1]	=	'startdate';
	arrFrm[1][2]	=	'enddate';
	arrFrm[1][3]	=	'reenddate';
	arrFrm[1][4]	=	'totalcnt';
	arrFrm[1][5]	=	'realcnt';
	arrFrm[1][6]	=	'filteringcnt';
	arrFrm[1][7]	=	'successcnt';
	arrFrm[1][8]	=	'failcnt';
	arrFrm[1][9]	=	'opencnt';
	arrFrm[1][10]	=	'noopencnt';

	var strCont = document.autofrm.testtxt.value;
	var tmpValue;

	strCont = strCont.replace(/\s{2,}/g,'\n'); 	//2ĭ�̻��� ������ "\n" ���� ����
	strCont = strCont.replace(/\n/g,'/');				//������ ������ "\n" �� "/" ���� ����

	var Wcount = strCont.length - strCont.replace(/\//g,'').length; // "/" �� ���� ����- ���� ������ ����

	var i = 0;

	while (i < Wcount){
		tmpValue=getTmpValue(strCont);
		strCont=getStrCont(strCont);

		tmpValue		= tmpValue.replace(/\s/g,'');			//����� ���ڿ��� ���� ����

		for(k=0;k<11;k++){  //tmpValue ���� ã�� ������ ���� ������ ����

				if(tmpValue.indexOf(arrFrm[0][k])==0){

						tmpValue=getTmpValue(strCont);
						strCont=getStrCont(strCont);

						var frm =eval('document.sform.' + arrFrm[1][k]);

						frm.value=tmpValue.replace(/[��](\W\S*)*/,'');
						i=i+1;
				}
		}
	i=i+1;
	}
	//���� ���ϱ�
	TnMailDataPercent();
}

// �Էµ� ������ ù "/" ���� ���峡���� ��ȯ
function getStrCont(strCont){

	var index	=	strCont.indexOf('/'); 					// "/"�� ��ġ�� ã�´�
	var len		= strCont.length;									// ��ü �����Ǳ��̸� ���Ѵ�
	strCont 	= strCont.substring(index+1,len);	// "/" ���� �κ� ���� ���峡���� ����

	return strCont;
}

// �Էµ� ������ ó������  ù "/"������ ���� ��ȯ
function getTmpValue(strCont){

	var index	=	strCont.indexOf('/'); 			// "/" �� ��ġ�� ã�´�
	tmpValue 	= strCont.substring(0,index);	// ��ü ���忡�� ó������ "/" ������ ���ڿ��� ����

	return tmpValue;
}

function TnMailDataPercent(){
	//�ǹ߼� ���		:
	document.sform.realpct.value = Math.round(eval(document.sform.realcnt.value/document.sform.totalcnt.value)*10000)/100;
	//���͸� ���		:
	document.sform.filteringpct.value = Math.round(eval(document.sform.filteringcnt.value/document.sform.totalcnt.value)*10000)/100;
	//���� �߼� ���:
	document.sform.successpct.value = Math.round(eval(document.sform.successcnt.value/document.sform.totalcnt.value)*10000)/100;
	//���� �߼� ���:
	document.sform.failpct.value = Math.round(eval(document.sform.failcnt.value/document.sform.totalcnt.value)*10000)/100;
	//�������			:
	document.sform.openpct.value = Math.round(eval(document.sform.opencnt.value/document.sform.totalcnt.value)*10000)/100;
	//�̿��� ���		:
	document.sform.noopenpct.value = Math.round(eval(document.sform.noopencnt.value/document.sform.totalcnt.value)*10000)/100;
}

</script>

<form name="autofrm" style="margin:0px;">
	<textarea name="testtxt" cols="40" rows="5"></textarea>
	<input type="button" value="����" onclick="autochk();" />
</form>

<% set omd = Nothing %>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->