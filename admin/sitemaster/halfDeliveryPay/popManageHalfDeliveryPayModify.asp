<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ��ۺ� �δ�ݾ� ���� �˾�
' Hieditor : 2020.08.27 ������ �߰�
'###########################################################
%>
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/classes/halfdeliverypay/halfdeliverypaycls.asp"-->
<%
Dim i, mode
Dim startdate, enddate, starttime, endtime
Dim idx
dim oHalfDeliveryView, loginUserId, dateModifyCheck

idx = requestCheckvar(request("idx"), 50)

loginUserId = session("ssBctId")

if Trim(idx) = "" then
	response.write "<script>alert('�������� ��η� �������ּ���.');window.close();</script>"
	response.end
end If

dateModifyCheck = false

'// halfdeliverypay View �����͸� �����´�.
set oHalfDeliveryView = new CgetHalfDeliveryPay
	oHalfDeliveryView.FRectIdx = idx
	oHalfDeliveryView.getHalfDeliveryPayview()


if Not(oHalfDeliveryView.FOneHalfDeliveryPay.Fstartdate="" or isNull(oHalfDeliveryView.FOneHalfDeliveryPay.Fstartdate)) Then
	starttime = Num2Str(hour(oHalfDeliveryView.FOneHalfDeliveryPay.Fstartdate),2,"0","R") &":"& Num2Str(minute(oHalfDeliveryView.FOneHalfDeliveryPay.Fstartdate),2,"0","R") &":"& Num2Str(second(oHalfDeliveryView.FOneHalfDeliveryPay.Fstartdate),2,"0","R")
else
	starttime = "00:00:00"
end if

if Not(oHalfDeliveryView.FOneHalfDeliveryPay.Fenddate="" or isNull(oHalfDeliveryView.FOneHalfDeliveryPay.Fenddate)) Then
	endtime = Num2Str(hour(oHalfDeliveryView.FOneHalfDeliveryPay.Fenddate),2,"0","R") &":"& Num2Str(minute(oHalfDeliveryView.FOneHalfDeliveryPay.Fenddate),2,"0","R") &":"& Num2Str(second(oHalfDeliveryView.FOneHalfDeliveryPay.Fenddate),2,"0","R")
else
	endtime = "00:00:00"
end if

'// �����Ϸ��� ��ǰ�� ���� ����ǰ� ������ ������ ������ �ȵ�
'// ���� ���� �� �̰ų� ����� ��ǰ�� ��쿡�� ������ ������ ����
If Cdate(left(now(), 10)) >= Cdate(left(oHalfDeliveryView.FOneHalfDeliveryPay.Fstartdate,10)) And Cdate(left(now(),10)) < Cdate(dateadd("d", 1, left(oHalfDeliveryView.FOneHalfDeliveryPay.Fenddate,10))) Then
	dateModifyCheck = false
Else
	dateModifyCheck = true
End If
%>
<link href="/js/jqueryui/css/jquery-ui.css" rel="stylesheet">
<link rel="stylesheet" type="text/css" href="/css/adminDefault.css" />
<link rel="stylesheet" type="text/css" href="/css/adminCommon.css" />
<style type="text/css">
html {overflow:auto;}
body {background-color:#fff;}
</style>
</head>
<body>
<script type="text/javascript" src="/js/jquery-1.10.1.min.js"></script>
<script type="text/javascript" src="/js/jquery-ui-1.10.3.custom.min.js"></script>
<script type="text/javascript" src="/js/jquery.swiper-3.3.1.min.js"></script>
<script type="text/javascript" src="/js/tag-it.min.js"></script>
<script type='text/javascript'>
document.domain = "10x10.co.kr";

function formatDate(date) { 
	var d = new Date(date), 
	month = '' + (d.getMonth() + 1), 
	day = '' + d.getDate(), 
	year = d.getFullYear(); 
	if (month.length < 2) month = '0' + month; 
	if (day.length < 2) day = '0' + day; 
	return [year, month, day].join('-'); 
}

function frmedit(){
	var frm  = document.frm;
	var today = new Date();

	<% If dateModifyCheck Then %>
		if(frm.startdate.value=="")
		{
			alert('�������� �Է��� �ּ���');
			frm.startdate.focus();
			return;
		}

		if(formatDate(today) >= formatDate(new Date(frm.startdate.value))) {
			alert('�������� ���� ���� ���� ���ں��� �����Ͻ� �� �ֽ��ϴ�.\n���� ������ ���ڷ� �Է����ּ���.');
			frm.startdate.focus();
			return;
		}
	<% End If %>

	if(frm.enddate.value=="")
	{
		alert('�������� �Է��� �ּ���');
		frm.enddate.focus();
		return;
	}

	if(formatDate(new Date(frm.enddate.value)) <= formatDate(new Date(frm.startdate.value))) {
		alert('�������� ������ ���ķθ� �����Ͻ� �� �ֽ��ϴ�.');
		frm.enddate.focus();
		return;
	}
	/*
	if(formatDate(today) >= formatDate(new Date(frm.startdate.value))) {
		alert('������ �������� ���� ���� ���� ���ں��� �����Ͻ� �� �ֽ��ϴ�.\n���� ������ ���ڷ� �Է����ּ���.');
		frm.enddate.focus();
		return;
	}
	*/


	if(frm.halfdeliverypay.value=="")
	{
		alert('��ۺ� �δ�ݾ��� �Է��� �ּ���');
		frm.halfdeliverypay.focus();
		return;
	}	

	if(!IsDigit(frm.halfdeliverypay.value)){
		alert("��ۺ� �δ�ݾ��� ���ڸ� �Է� �����մϴ�.");
		document.frm.halfdeliverypay.focus();
		return;
	}

	if(confirm("�����Ͻðڽ��ϱ�?")) {
		frm.submit();
	} else {
		return false;
	}
}

$(function()
{
	//�޷´�ȭâ ����
	var arrDayMin = ["��","��","ȭ","��","��","��","��"];
	var arrMonth = ["1��","2��","3��","4��","5��","6��","7��","8��","9��","10��","11��","12��"];
	$("#sDt").datepicker({
		dateFormat: "yy-mm-dd",
		prevText: '������', nextText: '������', yearSuffix: '��',
		dayNamesMin: arrDayMin,
		monthNames: arrMonth,
		showMonthAfterYear: true,
		numberOfMonths: 2,
		showCurrentAtPos: 1,
		showOn: "button",
		<% if idx<>"" then %>maxDate: "<%=oHalfDeliveryView.FOneHalfDeliveryPay.Fenddate%>",<% end if %>
		onClose: function( selectedDate ) {
			$( "#eDt" ).datepicker( "option", "minDate", selectedDate );
		}
	});
	$("#eDt").datepicker({
		dateFormat: "yy-mm-dd",
		prevText: '������', nextText: '������', yearSuffix: '��',
		dayNamesMin: arrDayMin,
		monthNames: arrMonth,
		showMonthAfterYear: true,
		numberOfMonths: 2,
		showOn: "button",
		<% if idx<>"" then %>minDate: "<%=oHalfDeliveryView.FOneHalfDeliveryPay.Fstartdate%>",<% end if %>
		onClose: function( selectedDate ) {
			$( "#sDt" ).datepicker( "option", "maxDate", selectedDate );
		}
	});
});

function checkLength(objname, maxlength)
{
	var objstr = objname.value;
	var objstrlen = objstr.length

	var maxlen = maxlength;
	var i = 0;
	var bytesize = 0;
	var strlen = 0;
	var onechar = "";
	var objstr2 = "";

	for (i = 0; i < objstrlen; i++)
	{
		onechar = objstr.charAt(i);

		if (escape(onechar).length > 4)
		{
			bytesize += 2;
		}
		else
		{
			bytesize++;
		}

		if (bytesize <= maxlen)
		{
			strlen = i + 1;
		}
	}

	if (bytesize > maxlen)
	{
		alert("���� ���ڿ��� �ʰ��Ͽ����ϴ�.\n�ѱ� ���� �ִ� "+maxlength/2+"�� ���� �ۼ��� �� �ֽ��ϴ�.");
		objstr2 = objstr.substr(0, strlen);
		objname.value = objstr2;
	}
	objname.focus();
}
</script>
<%' �˾� ������ : 750*800 %>
<form name="frm" method="post" action="halfdeliverypay_proc.asp">
<input type="hidden" name="mode" value="edit">
<input type="hidden" name="menupos" value="<%=menupos %>">
<input type="hidden" name="adminid" value="<%=loginUserId%>">
<input type="hidden" name="idx" value="<%=oHalfDeliveryView.FOneHalfDeliveryPay.Fidx%>">
<input type="hidden" name="defaultdeliveryType" value="<%=oHalfDeliveryView.FOneHalfDeliveryPay.FdefaultDeliveryType%>">
<input type="hidden" name="defaultFreeBeasongLimit" value="<%=oHalfDeliveryView.FOneHalfDeliveryPay.FdefaultFreeBeasongLimit%>">
<input type="hidden" name="defaultDeliverPay" value="<%=oHalfDeliveryView.FOneHalfDeliveryPay.FdefaultDeliverPay%>">

	<div class="popWinV17">
		<h1>����</h1>
		<div class="popContainerV17 pad30">
			<table class="tbType1 writeTb">
				<colgroup>
					<col width="18%" /><col width="" />
				</colgroup>
				<tbody>
				<tr>
					<th><div>��ȣ(idx) <strong class="cRd1"></strong></div></th>
					<td><%=oHalfDeliveryView.FOneHalfDeliveryPay.Fidx%></td>
				</tr>
				<% If dateModifyCheck Then %>
					<tr>
						<th><div>������ <strong class="cRd1">*</strong></div></th>
						<td>
							<input type="text" id="sDt" name="startdate" size="10" readonly value="<%=left(oHalfDeliveryView.FOneHalfDeliveryPay.Fstartdate,10)%>"/> <input type="hidden" name="starttime" value="<%=starttime%>" /><br/>
							<p class="tPad05 fs11 cGy1">- ������ ������ ������ ������ ���ķ� ������ �ȵǽŴٸ� �������� ���� �Է� �� �������� �Է����ּ���.</p>
						</td>
					</tr>
				<% Else %>
					<tr>
						<th><div>������ <strong class="cRd1">*</strong></div></th>
						<td>
							<%=left(oHalfDeliveryView.FOneHalfDeliveryPay.Fstartdate,10)%>
							<input type="hidden" name="startdate" value="<%=left(oHalfDeliveryView.FOneHalfDeliveryPay.Fstartdate,10)%>"/> <input type="hidden" name="starttime" value="<%=starttime%>" />
						</td>
					</tr>
				<% End If %>
				<tr>
					<th><div>������ <strong class="cRd1">*</strong></div></th>
					<td>
						<input type="text" id="eDt" name="enddate" size="10" readonly value="<%=left(oHalfDeliveryView.FOneHalfDeliveryPay.Fenddate,10)%>" /> <input type="hidden" name="endtime" value="<%=endtime%>" />
					</td>
				</tr>				
				<tr>
					<th><div>��ۺ� �δ�ݾ� <strong class="cRd1">*</strong></div></th>
					<td>
						<input type="text" id="halfdeliverypay" name="halfdeliverypay" size="10" value="<%=oHalfDeliveryView.FOneHalfDeliveryPay.FHalfDeliveryPay%>"/>
						<p class="tPad05 fs11 cGy1">- �޸����� ���ڸ� �־��ּ���.</p>
						<p class="tPad05 fs11 cRd1">- �ٹ� : ��ü�� �δ��ϴ� ��ۺ��Դϴ�.(��������)</p>
						<p class="tPad05 fs11 cRd1">- ���� : ��ü�� �����ϴ� ��ۺ��Դϴ�.(�߰�����)</p>						
					</td>
				</tr>
				<tr>
					<th><div>��뿩�� <strong class="cRd1">*</strong></div></th>
					<td>
						<span class="tPad05 col2">
							<label class="rMar20"><input type="radio" name="isusing" class="formRadio" value="N" <% If oHalfDeliveryView.FOneHalfDeliveryPay.Fisusing="N" Then %> checked <% End If %> /> ������</label>
							<label class="rMar20"><input type="radio" name="isusing" class="formRadio" value="Y" <% If oHalfDeliveryView.FOneHalfDeliveryPay.Fisusing="Y" Then %> checked <% End If %> /> �����</label>
						</span>
					</td>
				</tr>
				<tr>
					<th><div>��ϵ� ��ǰ</div></th>
					<td>
						<table class="tbType2 writeTb">
							<tr>
								<th align="center" width="12%">�̹���</th>
								<th align="center" width="11%">��ǰ�ڵ�</th>
								<th align="center" width="15%">�귣����̵�</th>
								<th align="center" width="28%">��ǰ��</th>
								<th align="center" width="13%">���ǹ�ۿ���</th>
								<th align="center" width="13%">�����۱��رݾ�</th>
								<th align="center" width="8%">��ۺ�</th>
							</tr>
							<tbody>
								<tr>
									<td width='11%'>
										<img src='<%=oHalfDeliveryView.FOneHalfDeliveryPay.Fsmallimage%>'>
									</td>
									<td width='10%'>
										<%=oHalfDeliveryView.FOneHalfDeliveryPay.Fitemid%>
									</td>
									<td width='14%'>
										<%=oHalfDeliveryView.FOneHalfDeliveryPay.Fbrandid%>
									</td>
									<td width='27%'>
										<%=oHalfDeliveryView.FOneHalfDeliveryPay.Fitemname%>
									</td>
									<td width='12%'>
										<%=getBeadalDivname(oHalfDeliveryView.FOneHalfDeliveryPay.FdefaultDeliveryType)%>
									</td>
									<td width='12%'>
										<%=oHalfDeliveryView.FOneHalfDeliveryPay.FdefaultFreeBeasongLimit%>
									</td>
									<td width='8%'>
										<%=oHalfDeliveryView.FOneHalfDeliveryPay.FdefaultDeliverPay%>
									</td>
								</tr>
							</tbody>
						</table>
					</td>
				</tr>									
				<tr>
					<th><div>�������</div></th>
					<td>
						<span class="tPad05 col2"><%=oHalfDeliveryView.FOneHalfDeliveryPay.Fadminid%>(<%=fnGetMyname(oHalfDeliveryView.FOneHalfDeliveryPay.Fadminid)%>)<br/><%=oHalfDeliveryView.FOneHalfDeliveryPay.Fregdate%></span>
					</td>
				</tr>
				<% If oHalfDeliveryView.FOneHalfDeliveryPay.Flastadminid <> "" Then %>
				<tr>
					<th><div>��������</div></th>
					<td>
						<span class="tPad05 col2 cRd1"><%=oHalfDeliveryView.FOneHalfDeliveryPay.Flastadminid%>(<%=fnGetMyname(oHalfDeliveryView.FOneHalfDeliveryPay.Flastadminid)%>)<br/><%=oHalfDeliveryView.FOneHalfDeliveryPay.Flastupdate%></span>
					</td>
				</tr>
				<% End If %>
				</tbody>
			</table>
		</div>
		<div class="popBtnWrap">
			<input type="button" value="���" onclick="window.close();" style="width:100px; height:30px;" />
			<input type="button" value="����" onclick="frmedit();" class="cRd1" style="width:100px; height:30px;" />
		</div>
	</div>
</form>
</body>
</html>
<%
	set oHalfDeliveryView = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
