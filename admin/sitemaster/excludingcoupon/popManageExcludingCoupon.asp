<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ���ʽ� ���� ���� ���� ��ǰor�귣�� ��� �˾�
' Hieditor : 2021.02.02 ������ �߰�
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

function frmExcludingCouponSubmit(){
	var frm  = document.frm;
	var today = new Date();

	if($("#mode").val() == "additem") {
		if (itemid == "") {
			alert('��ǰ�� ������ּ���.');
			return;
		}

		if(confirm("������ ����ϼ̴� ��ǰ�� ������쿣 �ش� ��ǰ�� ���/�������� �ʽ��ϴ�.\n��뿩��:"+$('input:radio[name=isusing]:checked').val()+"\n�ش� ������ ����Ͻðڽ��ϱ�?")) {
			frm.submit();
		} else {
			return false;
		}		
	}

	if($("#mode").val() == "addbrand") {
		if (makerid == "") {
			alert('�귣�带 ������ּ���.');
			return;
		}
		if(confirm("������ ����ϼ̴� �귣�尡 ������쿣 �ش� �귣��� ���/�������� �ʽ��ϴ�.\n��뿩��:"+$('input:radio[name=isusing]:checked').val()+"\n�ش� ������ ����Ͻðڽ��ϱ�?")) {
			frm.submit();
		} else {
			return false;
		}
	}
}


function jsAddItemData() {
	document.domain ="10x10.co.kr";
	var winAddItem;
	winAddItem = window.open('/common/pop_singleItemSelect.asp?target=frm&ptype=excludingcoupon','popAddItem','width=1000,height=600');
	winAddItem.focus();
}

function jsAddBrandData() {
	document.domain ="10x10.co.kr";
	var winAddItem;
	winAddItem = window.open('/admin/member/popBrandSearch.asp?frmName=frm&compName=makerid&isjsdomain=o','popAddBrand','width=1000,height=600');
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

function excludingCouponTypeChange(v) {
	if (v=='I') {
		$("#mode").val("additem");
		$("#itemAddArea").show();
		$("#brandAddArea").hide();
	} else if(v=="B") {
		$("#mode").val("addbrand");
		$("#itemAddArea").hide();
		$("#brandAddArea").show();
	}
}

</script>
<%' �˾� ������ : 750*800 %>
<form name="frm" method="post" action="excludingCoupon_proc.asp">
<input type="hidden" name="mode" id="mode" value="additem">
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
					<th><div>���� <strong class="cRd1">*</strong></div></th>
					<td>
						<span class="tPad05 col2">
							<label class="rMar20"><input type="radio" name="excludingCouponType" class="formRadio" value="I" checked onclick="excludingCouponTypeChange(this.value)"/> ��ǰ</label>
							<label class="rMar20"><input type="radio" name="excludingCouponType" class="formRadio" value="B" onclick="excludingCouponTypeChange(this.value)" /> �귣��</label>
						</span>
					</td>
				</tr>				
				<tr id="itemAddArea">
					<th><div>��ǰ��� <strong class="cRd1">*</strong></div></th>
					<td>
						<input type="text" id="itemid" name="itemid" size="10" />
						<input type="button" value="��ǰ�˻�" onclick="jsAddItemData();" style="width:100px;" />
					</td>
				</tr>				
				<tr id="brandAddArea" style="display:none;">
					<th><div>�귣�� ��� <strong class="cRd1">*</strong></div></th>
					<td>
						<input type="text" id="makerid" name="makerid" size="10" />
						<input type="button" value="�귣�� �˻�" onclick="jsAddBrandData();" style="width:100px;" />
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
				</tbody>
			</table>
		</div>
		<div class="popBtnWrap">
			<input type="button" value="���" onclick="window.close();" style="width:100px; height:30px;" />
			<input type="button" value="���" onclick="frmExcludingCouponSubmit();" class="cRd1" style="width:100px; height:30px;" />
		</div>
	</div>
</form>
</body>
</html>
<%
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
