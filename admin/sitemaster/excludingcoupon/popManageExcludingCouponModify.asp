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
<!-- #include virtual="/lib/classes/excludingcoupon/excludingcouponcls.asp"-->
<%
Dim i, mode
Dim idx
dim oExcludingCouponView, loginUserId

idx = requestCheckvar(request("idx"), 50)

loginUserId = session("ssBctId")

if Trim(idx) = "" then
	response.write "<script>alert('�������� ��η� �������ּ���.');window.close();</script>"
	response.end
end If

'// halfdeliverypay View �����͸� �����´�.
set oExcludingCouponView = new CgetExcludingCoupon
	oExcludingCouponView.FRectIdx = idx
	oExcludingCouponView.getExcludingCouponview()
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

function frmedit(){
	var frm  = document.frm;

	if(confirm("�����Ͻðڽ��ϱ�?")) {
		frm.submit();
	} else {
		return false;
	}
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
</script>
<%' �˾� ������ : 750*800 %>
<form name="frm" method="post" action="excludingCoupon_proc.asp">
<input type="hidden" name="mode" value="edit">
<input type="hidden" name="menupos" value="<%=menupos %>">
<input type="hidden" name="adminid" value="<%=loginUserId%>">
<input type="hidden" name="idx" value="<%=oExcludingCouponView.FOneExcludingCoupon.Fidx%>">
<input type="hidden" name="excludingCouponType" value="<%=oExcludingCouponView.FOneExcludingCoupon.Ftype%>">
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
					<td><%=oExcludingCouponView.FOneExcludingCoupon.Fidx%></td>
				</tr>
				<tr>
					<th><div>���� <strong class="cRd1">*</strong></div></th>
					<td>
                        <%
                            If oExcludingCouponView.FOneExcludingCoupon.Ftype = "I" Then
                                response.write "��ǰ"
                            ElseIf oExcludingCouponView.FOneExcludingCoupon.Ftype = "B" Then
                                response.write "�귣��"
                            End If
                        %>
					</td>
				</tr>
                <% If oExcludingCouponView.FOneExcludingCoupon.Ftype = "I" Then %>
                    <tr id="itemAddArea">
                        <th><div>��ǰ��� <strong class="cRd1">*</strong></div></th>
                        <td>
                            <input type="text" id="itemid" name="itemid" size="10"  value="<%=oExcludingCouponView.FOneExcludingCoupon.FItemID%>" />
                            <input type="button" value="��ǰ�˻�" onclick="jsAddItemData();" style="width:100px;" />
                        </td>
                    </tr>
                <% ElseIf oExcludingCouponView.FOneExcludingCoupon.Ftype = "B" Then %>
                    <tr id="brandAddArea">
                        <th><div>�귣�� ��� <strong class="cRd1">*</strong></div></th>
                        <td>
                            <input type="text" id="makerid" name="makerid" size="10" value="<%=oExcludingCouponView.FOneExcludingCoupon.Fbrandid%>" />
                            <input type="button" value="�귣�� �˻�" onclick="jsAddBrandData();" style="width:100px;" />
                        </td>
                    </tr>
                <% End If %>
				<tr>
					<th><div>��뿩�� <strong class="cRd1">*</strong></div></th>
					<td>
						<span class="tPad05 col2">
							<label class="rMar20"><input type="radio" name="isusing" class="formRadio" value="N" <% If oExcludingCouponView.FOneExcludingCoupon.Fisusing="N" Then %> checked <% End If %> /> ������</label>
							<label class="rMar20"><input type="radio" name="isusing" class="formRadio" value="Y" <% If oExcludingCouponView.FOneExcludingCoupon.Fisusing="Y" Then %> checked <% End If %> /> �����</label>
						</span>
					</td>
				</tr>
				<tr>
					<th><div>�������</div></th>
					<td>
						<span class="tPad05 col2"><%=oExcludingCouponView.FOneExcludingCoupon.Fadminid%>(<%=fnGetMyname(oExcludingCouponView.FOneExcludingCoupon.Fadminid)%>)<br/><%=oExcludingCouponView.FOneExcludingCoupon.Fregdate%></span>
					</td>
				</tr>
				<% If oExcludingCouponView.FOneExcludingCoupon.Flastadminid <> "" Then %>
				<tr>
					<th><div>��������</div></th>
					<td>
						<span class="tPad05 col2 cRd1"><%=oExcludingCouponView.FOneExcludingCoupon.Flastadminid%>(<%=fnGetMyname(oExcludingCouponView.FOneExcludingCoupon.Flastadminid)%>)<br/><%=oExcludingCouponView.FOneExcludingCoupon.Flastupdate%></span>
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
	set oExcludingCouponView = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
