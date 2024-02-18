<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ��ۺ� �δ�ݾ� ��� �˾�
' Hieditor : 2020.08.27 ������ �߰�
'###########################################################
%>
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
Dim i, loginUserId
loginUserId = session("ssBctId")
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

function frmHalfDeliveryPaySubmit(){
	var frm  = document.frm;
	var today = new Date();

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

	if(frm.enddate.value=="")
	{
		alert('�������� �Է��� �ּ���');
		frm.enddate.focus();
		return;
	}

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

	if(typeof frm.iid == "undefined")
	{
		alert('��ǰ�� ������ּ���.');
		return;
	}

	if(confirm("�Է��Ͻ� ������,������,��ۺ�δ�ݾ�,��뿩�δ�\n���â���� ����Ͻ� ��ǰ�� �ϰ����� �˴ϴ�.\n������ ����ϼ̴� ��ǰ�� ������쿣 �ش� ��ǰ�� ���/�������� �ʽ��ϴ�.\n\n������:"+frm.startdate.value+" "+frm.starttime.value+"\n������:"+frm.enddate.value+" "+frm.endtime.value+"\n��ۺ� �δ�ݾ�:"+frm.halfdeliverypay.value+"��\n��뿩��:"+$('input:radio[name=isusing]:checked').val()+"\n�ش� ������ ����Ͻðڽ��ϱ�?")) {
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
	var today = new Date();
	$("#sDt").datepicker({
		dateFormat: "yy-mm-dd",
		prevText: '������', nextText: '������', yearSuffix: '��',
		dayNamesMin: arrDayMin,
		monthNames: arrMonth,
		showMonthAfterYear: true,
		numberOfMonths: 2,
		showCurrentAtPos: 1,
		showOn: "button",
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
		onClose: function( selectedDate ) {
			$( "#sDt" ).datepicker( "option", "maxDate", selectedDate );
		}
	});
});

function jsAddItemData() {
	document.domain ="10x10.co.kr";
	var winAddItem;
	winAddItem = window.open('pop_additemlist.asp','popAddItem','width=1000,height=600');
	winAddItem.focus();
}

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

function userAreaDeleteItem(trid) {
	var tr = $(trid).parent().parent();
	tr.remove();
}

function viewUserAddItemListData() {
	var str_array = $("#viewitemdataparent").val().split(',');
	var str_array_detail;
	for(var i = 0; i < str_array.length; i++) {
		str_array_detail = str_array[i].split('|');
		if($("#itemListArea").html().indexOf("viewdata"+str_array_detail[1]) == -1) {
			$("#itemListArea").append("<tr id=viewdata"+str_array_detail[1]+"><td width='11%'><img src='"+str_array_detail[0]+"'></td><td width='10%'>"+str_array_detail[1]+"</td><td width='14%'>"+str_array_detail[2]+"</td><td width='27%'>"+str_array_detail[3]+"</td><td width='12%'>"+str_array_detail[4]+"</td><td width='12%'>"+str_array_detail[5]+"</td><td width='8%'>"+str_array_detail[6]+"</td><td width='10%'><button onclick='userAreaDeleteItem(this);'>����</button></td><input type='hidden' name='iid' value='"+str_array_detail[1]+"'>");
		}
	}
}

</script>
<%' �˾� ������ : 750*800 %>
<form name="frm" method="post" action="halfdeliverypay_proc.asp">
<input type="hidden" name="mode" value="add">
<input type="hidden" name="menupos" value="<%=menupos %>">
<input type="hidden" name="adminid" value="<%=loginUserId%>">
<input type="hidden" name="viewitemdataparent" id="viewitemdataparent">
	<div class="popWinV17">
		<h1>���</h1>
		<div class="popContainerV17 pad30">
			<table class="tbType1 writeTb">
				<colgroup>
					<col width="18%" /><col width="" />
				</colgroup>
				<tbody>
				<tr>
					<th><div>������ <strong class="cRd1">*</strong></div></th>
					<td>
						<input type="text" id="sDt" name="startdate" size="10" readonly /> <input type="hidden" name="starttime" value="00:00:00" />
					</td>
				</tr>
				<tr>
					<th><div>������ <strong class="cRd1">*</strong></div></th>
					<td>
						<input type="text" id="eDt" name="enddate" size="10" readonly /> <input type="hidden" name="endtime" value="23:59:59" />
					</td>
				</tr>
				<tr>
					<th><div>��ۺ� �δ�ݾ� <strong class="cRd1">*</strong></div></th>
					<td>
						<input type="text" id="halfdeliverypay" name="halfdeliverypay" size="10" />
						<p class="tPad05 fs11 cGy1">- �޸����� ���ڸ� �־��ּ���.</p>
						<p class="tPad05 fs11 cRd1">- �ٹ� : ��ü�� �δ��ϴ� ��ۺ��Դϴ�.(��������)</p>
						<p class="tPad05 fs11 cRd1">- ���� : ��ü�� �����ϴ� ��ۺ��Դϴ�.(�߰�����)</p>
					</td>
				</tr>
				<tr>
					<th><div>��뿩�� <strong class="cRd1">*</strong></div></th>
					<td>
						<span class="tPad05 col2">
							<label class="rMar20"><input type="radio" name="isusing" class="formRadio" value="N" checked /> ������</label>
							<label class="rMar20"><input type="radio" name="isusing" class="formRadio" value="Y" /> �����</label>
						</span>
					</td>
				</tr>
				<tr>
					<th><div>��ǰ��� <strong class="cRd1">*</strong></div></th>
					<td>
						<input type="button" value="��ǰ���" onclick="jsAddItemData();" style="width:100px; height:30px;" />
						<p>&nbsp;</p>
						<table class="tbType2 writeTb">
							<tr>
								<th align="center" width="12%">�̹���</th>
								<th align="center" width="11%">��ǰ�ڵ�</th>
								<th align="center" width="15%">�귣����̵�</th>
								<th align="center" width="28%">��ǰ��</th>
								<th align="center" width="13%">���ǹ�ۿ���</th>
								<th align="center" width="13%">�����۱��رݾ�</th>
								<th align="center" width="8%">��ۺ�</th>
								<th width="10%"></th>
							</tr>
							<tbody id="itemListArea">
							</tbody>
						</table>
					</td>
				</tr>				
				</tbody>
			</table>
		</div>
		<div class="popBtnWrap">
			<input type="button" value="���" onclick="window.close();" style="width:100px; height:30px;" />
			<input type="button" value="���" onclick="frmHalfDeliveryPaySubmit();" class="cRd1" style="width:100px; height:30px;" />
		</div>
	</div>
</form>
</body>
</html>
<%
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
