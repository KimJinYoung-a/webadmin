<%@ language=vbscript %>
<% option explicit %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'####################################################
' Description :  �������� ����ٹ�����
' History : 2011.03.17 �ѿ�� ����
'           2012.02.15 ������- �̴ϴ޷� ��ü
'####################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/common/incSessionAdminorShop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/staff/staff_cls.asp"-->
<%
dim idx, sDt, sTm, eDt, eTm ,oAgitCal ,shopid ,empno
dim userid,username,posit_sn,part_sn,ChkStart,ChkEnd,etcComment
	idx = request("idx")
	shopid = request("shopid")

	if shopid = "" then
		resposne.write "<script>alert('������ �������� �ʾҽ��ϴ�'); self.close(); </script>"
		dbget.close()	:	response.end
	end if
	
'// ���� ����
if idx<>"" then
	
	Set oAgitCal = new CAgitCalendar
		oAgitCal.frectidx = idx
		oAgitCal.read()
		
		if oAgitCal.ftotalcount >0 then
			userid			= oAgitCal.FOneItem.Fuserid
			username		= oAgitCal.FOneItem.Fusername
			posit_sn		= oAgitCal.FOneItem.Fposit_sn
			part_sn			= oAgitCal.FOneItem.Fpart_sn						
			ChkStart		= oAgitCal.FOneItem.FChkStart
			ChkEnd			= oAgitCal.FOneItem.FChkEnd			
			etcComment		= oAgitCal.FOneItem.FetcComment
			shopid			= oAgitCal.FOneItem.fshopid
			empno 			= oAgitCal.FOneItem.fempno
		
			sDt = left(ChkStart,10)
			eDt = left(ChkEnd,10)
			sTm = Num2Str(Hour(ChkStart),2,"0","R") & ":" & Num2Str(Minute(ChkStart),2,"0","R")& ":" & Num2Str(second(ChkStart),2,"0","R")
			eTm = Num2Str(Hour(ChkEnd),2,"0","R") & ":" & Num2Str(Minute(ChkEnd),2,"0","R")& ":" & Num2Str(second(ChkEnd),2,"0","R")
		end if
	Set oAgitCal = Nothing
end if

if sDt="" then sDt=date
if sTm="" then sTm="00:00:00"
if eDt="" then eDt=dateAdd("d",1,date)
if eTm="" then eTm="24:00:00"
	
part_sn = "18"	
%>

<script language="javascript1.2" type="text/javascript" src="/js/datetime.js"></script>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />

<script language="javascript">

	// ��� Ȯ�� �� ó��
	function chk_form(form)	{
		if(form.shopid.value=="") {
			alert("������ �������ּ���");
			form.shopid.focus();
			return false;
		}

		if(form.chkCfm.value!="Y") {
			alert("�˻� ��ư�� ���� ������ �˻��� �ּ���.");
			form.SearchText.focus();
			return false;
		}

		if(form.empno.value == "") {
			alert("�����ȣ�� �Է� �ϼ���");
			form.empno.focus();
			return false;
		}

		if(form.posit_sn.value == "") {
			alert("������ �������ּ���.");
			form.posit_sn.focus();
			return false;
		}

		if(form.username.value == "") {
			alert("�̸��� �Է����ּ���.");
			form.username.focus();
			return false;
		}

		if(form.part_sn.value == "") {
			alert("�ҼӺμ��� �������ּ���.");
			form.part_sn.focus();
			return false;
		}

		//if(getDayInterval(toDate(form.ChkStart.value), toDate('<%=date%>'))>0) {
		//	alert("������ ��¥�� ����Ͻ� �� �����ϴ�. ��¥�� Ȯ�����ּ���.");
		//	return false;
		//}

		if(getDayInterval(toDate(form.ChkStart.value), toDate(form.ChkEnd.value))<0) {
			alert("�Ⱓ�� �߸��Ǿ� �ֽ��ϴ�. ��¥�� Ȯ�����ּ���.");
			return false;
		}

		if(confirm(form.ChkStart.value +"~"+form.ChkEnd.value +"�Ⱓ��("+form.uTerm.value+"��) " + form.username.value + "���� ����Ͻðڽ��ϱ�?"))	{
		return true;
		}
		return false;
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
	
		frm.uTerm.value = getDayInterval(startdate, enddate)+1;
	}

	//���� ���̵� �˻� �� ���ó��� �ڵ��Է�
	function chkTenMember() {
		var SearchType;
		var SearchText;
		var shopid;
		
		if(frm.SearchType.value == '') {
			alert("�˻��Ͻ� ������ �����ϼ���");
			frm.SearchType.focus();
			return;
		}

		if(frm.SearchText.value == '') {
			alert("�˻��Ͻ� ���� �Է� �ϼ���");
			frm.SearchText.focus();
			return;
		}		

		if(frm.shopid.value == '') {
			alert("���õ� ������ �����ϴ�");			
			return;
		}	
		
		SearchType = frm.SearchType.value;
		SearchText = frm.SearchText.value;
		shopid = frm.shopid.value;
		document.getElementById("ifmProc").src="/common/offshop/staff/actionTenUser.asp?SearchType="+SearchType+"&SearchText="+SearchText+"&shopid="+shopid;
	}

	//��üó��
	function delBook() {
		if(confirm("�� ���೻���� �����Ͻðڽ��ϱ�?"))	{
			frm.mode.value = "del";
			frm.submit();
		}
	}

</script>

<table width="100%" border="0" cellpadding="5" cellspacing="0" class="a">
<form name="frm" method="POST" action="/common/offshop/staff/shop_staff_schedule_Process.asp" onsubmit="return chk_form(this)">
<input type="hidden" name="mode" value="<%=chkIIF(idx="","add","modi")%>">
<input type="hidden" name="idx" value="<%=idx%>">
<tr>
	<td>
		<table width="100%" border="0" cellpadding="5" cellspacing="1" class="a" bgcolor="#909090">
		<tr bgcolor="#FFFFFF">
			<td width="120" bgcolor="<%=adminColor("sky")%>" align="center"><b>����</b></td>
			<td>				
				<%= shopid %><input type="hidden" name="shopid" value="<%=shopid%>">
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td width="120" bgcolor="<%=adminColor("sky")%>" align="center"><b>����˻�</b></td>
			<td>
				<select name="SearchType">
					<option value="2">�̸�</option>
					<option value="1">���̵�</option>					
					<option value="3">���</option>
				</select>				
				<input type="text" name="SearchText" size="20" class="text">
				<input type="button" value="�˻�" class="button_s" style="width:55px;text-align:center;" onclick="chkTenMember()">				
				<input type="hidden" name="chkCfm" value="<%=chkIIF(idx="","N","Y")%>">
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td width="120" bgcolor="<%=adminColor("sky")%>" align="center"><b>�����ȣ</b></td>
			<td>				
				<input type="text" name="empno" size="16" class="text" value="<%=empno%>">
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td width="120" bgcolor="<%=adminColor("sky")%>" align="center"><b>���̵�</b></td>
			<td>				
				<input type="text" name="userid" size="16" class="text" value="<%=userid%>">
			</td>
		</tr>		
		<tr bgcolor="#FFFFFF">
			<td width="120" bgcolor="<%=adminColor("sky")%>" align="center"><b>����/�̸�</b></td>
			<td>
				<%=printPositOption("posit_sn", posit_sn)%>
				<input type="text" name="username" size="16" class="text" value="<%=username%>">
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td width="120" bgcolor="<%=adminColor("sky")%>" align="center"><b>�ҼӺμ�</b></td>
			<td><%=printPartOption("part_sn", part_sn)%></td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td width="120" bgcolor="<%=adminColor("sky")%>" align="center"><b>�Ⱓ</b></td>
			<td style="line-height:18px;">
				<input id="ChkStart" name="ChkStart" value="<%=sDt%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="ChkStart_trigger" border="0" style="cursor:pointer" align="absmiddle" />
		    	<input type="text" name="ChkSTime" size="8" maxlength="8" class="text" value="<%=sTm%>">
		    	~
				<input id="ChkEnd" name="ChkEnd" value="<%=eDt%>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="ChkEnd_trigger" border="0" style="cursor:pointer" align="absmiddle" />
		    	<input type="text" name="ChkETime" size="8" maxlength="8" class="text" value="<%=eTm%>">
		    	<font color=gray>(<input type="text" name="uTerm" readonly class="text" value="<%=DateDiff("d",sDt,eDt)+1%>" style="text-align:right; width:20px; border:0px; color:gray;">��)</font>
				<script language="javascript">
					var CAL_Start = new Calendar({
						inputField : "ChkStart", trigger    : "ChkStart_trigger",
						onSelect: function() {
							var date = Calendar.intToDate(this.selection.get());
							CAL_End.args.min = date;
							CAL_End.redraw();
							this.hide();
						}, bottomBar: true, dateFormat: "%Y-%m-%d"
					});
					var CAL_End = new Calendar({
						inputField : "ChkEnd", trigger    : "ChkEnd_trigger",
						onSelect: function() {
							var date = Calendar.intToDate(this.selection.get());
							CAL_Start.args.max = date;
							CAL_Start.redraw();
							this.hide();
						}, bottomBar: true, dateFormat: "%Y-%m-%d"
					});
				</script>
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td width="120" bgcolor="<%=adminColor("sky")%>" align="center"><b>���</b></td>
			<td><textarea name="etcComment" class="textarea" style="width:100%; height:50px;"><%=etcComment%></textarea></td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td colspan="2" align="center">
				<input type="submit" value="�� ��" class="button" style="width:60px;text-align:center;">
				<% if idx<>"" then %>
					<input type="button" value="�� ��" class="button" style="width:60px;text-align:center;" onclick="delBook()">
				<% end if %>
			</td>
		</tr>
		</table>
	</td>
</tr>
</form>
</table>

<iframe id="ifmProc" src="" width=0 height=0 frameborder="0"></iframe>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->