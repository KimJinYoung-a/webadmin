<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/sitemasterclass/100proshopCls.asp" -->
<%
dim eCode,idx,mode
eCode = request("eC")
idx = request("idx")
mode = request("mode")
%>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script language='javascript'>
function SubmitForm(){
	if (document.SubmitFrm.itemid.value.length < 1){
		alert('��ǰ��ȣ �� �Է��ϼ���');
		document.SubmitFrm.itemid.focus();
		return;
	}
	if (document.SubmitFrm.startdate.value.length < 1){
		alert('�������� �Է��ϼ���');
		document.SubmitFrm.startdate.focus();
		return;
	}
	if (document.SubmitFrm.enddate.value.length < 1){
		alert('�������� �Է��ϼ���');
		document.SubmitFrm.enddate.focus();
		return;
	}


	if (document.SubmitFrm.couponname.value.length < 1){
		alert('�������� �Է��ϼ���.');
		document.SubmitFrm.couponname.focus();
		return;
	}

	if (document.SubmitFrm.couponvalue.value.length < 1){
		alert('�����ݾ��� �Է��ϼ���.');
		document.SubmitFrm.couponvalue.focus();
		return;
	}

	if (document.SubmitFrm.couponstartdate.value.length < 1){
		alert('������ȿ�Ⱓ �������� �Է��ϼ���.');
		document.SubmitFrm.couponstartdate.focus();
		return;
	}

	if (document.SubmitFrm.couponexpiredate.value.length < 1){
		alert('������ȿ�Ⱓ �������� �Է��ϼ���.');
		document.SubmitFrm.couponexpiredate.focus();
		return;
	}

	if (document.SubmitFrm.minbuyprice.value.length < 1){
		alert('�ּ� ���űݾ��� �Է��ϼ���.');
		document.SubmitFrm.minbuyprice.focus();
		return;
	}

	var ret = confirm('���� �Ͻðڽ��ϱ�?');
	if (ret) {
		document.SubmitFrm.submit();
	}
}

function calender_open(objectname) {
       document.all.cal.style.display="";
	   document.all.cal.style.left = event.offsetX;
	   document.all.cal.style.top = event.offsetY + 200;
	   document.SubmitFrm.objname.value = objectname;

//	   alert("X-��ǥ : " + event.offsetX + "\n" + "Y-��ǥ : " + event.offsetY);
}

function getIteminfo(idx) {
	window.open('100proshop_Item_get.asp?eC=' + idx,'getwin','width=350,height=300,resizable=yes,scrollbars=yes,');
}

</script>
<br><br>


<table width="700" border="1" cellpadding="0" cellspacing="0" class="a" bordercolordark="White" bordercolorlight="black" align="center">
  <form name="SubmitFrm" method="post" action="do100proshopitem.asp" onsubmit="return false;" >
    <input type="hidden" name="mode" value="<% = mode %>">
	<input type="hidden" name="coupontype" value="2">
	<input type="hidden" name="eCode" value="<% = eCode %>">
	<input type="hidden" name="idx" value="<% = idx %>">
<%
if mode = "modify" then
dim o100pro
set o100pro = new C100ProShop
o100pro.FCurrPage = 1
o100pro.FPageSize = 1
o100pro.read idx
%>
	<tr>
	  <td width="100">100%�޹�ȣ</td>
	  <td><% = eCode %>
	  &nbsp;<input type="button" value="�ҷ�����" onclick="getIteminfo('<%= eCode %>')" /></td>
	</tr>
	<tr>
	  <td width="100">��ǰ��ȣ</td>
	  <td><input type="text" name="itemid" size="10" value="<% = o100pro.FItemList(0).Fitemid %>"></td>
	</tr>
	<tr>
	  <td width="100">�����߱ޱ�����</td>
	  <td>
		<input id="startdate" name="startdate" value="<% = FormatDateTime(o100pro.FItemList(0).FStartDate,2) %>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="startdate_trigger" border="0" style="cursor:pointer" align="absmiddle" /> ~
		<input id="enddate" name="enddate" value="<% = FormatDateTime(o100pro.FItemList(0).FEndDate,2) %>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="enddate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
		<script language="javascript">
			var CAL_Start = new Calendar({
				inputField : "startdate", trigger    : "startdate_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					CAL_End.args.min = date;
					CAL_End.redraw();
					this.hide();
				}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
			var CAL_End = new Calendar({
				inputField : "enddate", trigger    : "enddate_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					CAL_Start.args.max = date;
					CAL_Start.redraw();
					this.hide();
				}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
		</script>
		(YYYY-MM-DD) <br> * ���ÿ��ΰ� N�ΰ�쿡�� �����߱ޱ����� ���� �ֹ��� ��� ������ �߱޵�
	  </td>
	</tr>
	<tr>
	  <td width="100">���ÿ���</td>
	  <td>
	  	<input type="radio" name="isusing" value="Y" <% if o100pro.FItemList(0).FIsUsing="Y" then response.write "checked" %> >Y
	  	<input type="radio" name="isusing" value="N" <% if o100pro.FItemList(0).FIsUsing="N" then response.write "checked" %> >N
	  	<br>(���ÿ��θ� N�� ������ ��� ���� �߱� �������� ������Ϸ� �ؾ���)
	  </td>
	</tr>
	<tr>
	  <td width="100" colspan="2" bgcolor="#EEEEEE">��������</td>
	</tr>
	<tr>
	  <td width="100">������</td>
	  <td>
	  	<input type="text" name="couponname" value="<%= o100pro.FItemList(0).FCouponName %>" size="40" class="input_b">
	  	(ex: 100% �� ���� 0,000��)
	  </td>
	</tr>
	<tr>
	  <td width="100">�����ݾ�</td>
	  <td>
	  	<input type="text" name="couponvalue" value="<%= o100pro.FItemList(0).FCouponValue %>" size="6" class="input_b">��
	  </td>
	</tr>
	<tr>
	  <td width="100">������ȿ�Ⱓ</td>
	  <td>
		<input id="couponstartdate" name="couponstartdate" value="<% = FormatDateTime(o100pro.FItemList(0).FCouponStartDate,2) %>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="couponstartdate_trigger" border="0" style="cursor:pointer" align="absmiddle" /> ~
		<input id="couponexpiredate" name="couponexpiredate" value="<% = FormatDateTime(o100pro.FItemList(0).FCouponExpireDate,2) %>" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="couponexpiredate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
		<script language="javascript">
			var CPN_Start = new Calendar({
				inputField : "couponstartdate", trigger    : "couponstartdate_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					CPN_End.args.min = date;
					CPN_End.redraw();
					this.hide();
				}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
			var CPN_End = new Calendar({
				inputField : "couponexpiredate", trigger    : "couponexpiredate_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					CPN_Start.args.max = date;
					CPN_Start.redraw();
					this.hide();
				}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
		</script>
		(YYYY-MM-DD) : �����߱ޱⰣ �� ��� ��ȿ�Ⱓ
	  </td>
	</tr>
	<tr>
	  <td width="100">�ּұ��űݾ�</td>
	  <td>
	  	<input type="text" name="minbuyprice" value="<%= o100pro.FItemList(0).Fminbuyprice %>" size="6" class="input_b">��
	  </td>
	</tr>
<!-- // 2009�� ����Ʈ ������
	<tr>
	  <td width="100" colspan="2" bgcolor="#EEEEEE">MD�ڸ�Ʈ ����</td>
	</tr>
	<tr>
	  <td width="100">�ڸ�Ʈ1</td>
	  <td>
	  	MD��<input type="text" name="mdname1" value="<%= o100pro.FItemList(0).Fmdname1 %>" size="16" class="input_b"><br>
	  	<textarea name="mdcomment1" cols="60" rows="5"><%= o100pro.FItemList(0).Fmdcomment1 %></textarea>
	  </td>
	</tr>
	<tr>
	  <td width="100">�ڸ�Ʈ2</td>
	  <td>
	  	MD��<input type="text" name="mdname2" value="<%= o100pro.FItemList(0).Fmdname2 %>" size="16" class="input_b"><br>
	  	<textarea name="mdcomment2" cols="60" rows="5"><%= o100pro.FItemList(0).Fmdcomment2 %></textarea>
	  </td>
	</tr>
	<tr>
	  <td width="100">�ڸ�Ʈ3</td>
	  <td>
	  	MD��<input type="text" name="mdname3" value="<%= o100pro.FItemList(0).Fmdname3 %>" size="16" class="input_b"><br>
	  	<textarea name="mdcomment3" cols="60" rows="5"><%= o100pro.FItemList(0).Fmdcomment3 %></textarea>
	  </td>
	</tr>
// -->
	<tr>
	  <td colspan="2" align="center">
	  	<input type="button" value="����" onClick="SubmitForm();">
	  </td>
	</tr>
	</form>
</table>
<%
set o100pro = Nothing
%>
<%
else
%>
	<tr>
	  <td width="100">100%�޹�ȣ</td>
	  <td><% = eCode %>&nbsp;<input type="button" value="�ҷ�����" style="font:12px;width:70px;height:20px;" onclick="getIteminfo('<%= eCode %>')" /></td>
	</tr>
	<tr>
	  <td width="100">��ǰ��ȣ</td>
	  <td><input type="text" name="itemid" size="10" class="input_b" ></td>
	</tr>
	<tr>
	  <td width="100">�����߱ޱ�����</td>
	  <td>
		<input id="startdate" name="startdate" value="" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="startdate_trigger" border="0" style="cursor:pointer" align="absmiddle" /> ~
		<input id="enddate" name="enddate" value="" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="enddate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
		<script language="javascript">
			var CAL_Start = new Calendar({
				inputField : "startdate", trigger    : "startdate_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					CAL_End.args.min = date;
					CAL_End.redraw();
					this.hide();
				}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
			var CAL_End = new Calendar({
				inputField : "enddate", trigger    : "enddate_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					CAL_Start.args.max = date;
					CAL_Start.redraw();
					this.hide();
				}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
		</script>
		(YYYY-MM-DD) <br> * ���ÿ��ΰ� N�ΰ�쿡�� �����߱ޱ����� ���� �ֹ��� ��� ������ �߱޵�
	  </td>
	</tr>
	<tr>
	  <td width="100">���ÿ���</td>
	  <td>
	  	<input type="radio" name="isusing" value="Y" checked>Y
	  	<input type="radio" name="isusing" value="N">N
	  	<br>(���ÿ��θ� N�� ������ ��� ���� �߱� �������� ������Ϸ� �ؾ���)
	  </td>
	</tr>
	<tr>
	  <td width="100" colspan="2" bgcolor="#EEEEEE">��������</td>
	</tr>
	<tr>
	  <td width="100">������</td>
	  <td>
	  	<input type="text" name="couponname" value="" size="40" class="input_b">
	  	(ex: 100% �� ���� 0,000��)
	  </td>
	</tr>
	<tr>
	  <td width="100">�����ݾ�</td>
	  <td>
	  	<input type="text" name="couponvalue" value="" size="6" class="input_b">��
	  </td>
	</tr>
	<tr>
	  <td width="100">������ȿ�Ⱓ</td>
	  <td>
		<input id="couponstartdate" name="couponstartdate" value="" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="couponstartdate_trigger" border="0" style="cursor:pointer" align="absmiddle" /> ~
		<input id="couponexpiredate" name="couponexpiredate" value="" class="text" size="10" maxlength="10" /><img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="couponexpiredate_trigger" border="0" style="cursor:pointer" align="absmiddle" />
		<script language="javascript">
			var CPN_Start = new Calendar({
				inputField : "couponstartdate", trigger    : "couponstartdate_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					CPN_End.args.min = date;
					CPN_End.redraw();
					this.hide();
				}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
			var CPN_End = new Calendar({
				inputField : "couponexpiredate", trigger    : "couponexpiredate_trigger",
				onSelect: function() {
					var date = Calendar.intToDate(this.selection.get());
					CPN_Start.args.max = date;
					CPN_Start.redraw();
					this.hide();
				}, bottomBar: true, dateFormat: "%Y-%m-%d"
			});
		</script>
		(YYYY-MM-DD) : �����߱ޱⰣ �� ��� ��ȿ�Ⱓ
	  </td>
	</tr>
	<tr>
	  <td width="100">�ּұ��űݾ�</td>
	  <td>
	  	<input type="text" name="minbuyprice" value="" size="6" class="input_b">��
	  </td>
	</tr>
<!-- // 2009�� ����Ʈ ������
	<tr>
	  <td width="100" colspan="2" bgcolor="#EEEEEE">MD�ڸ�Ʈ ����</td>
	</tr>
	<tr>
	  <td width="100">�ڸ�Ʈ1</td>
	  <td>
	  	MD��<input type="text" name="mdname1" value="" size="16" class="input_b"><br>
	  	<textarea name="mdcomment1" cols="60" rows="5"></textarea>
	  </td>
	</tr>
	<tr>
	  <td width="100">�ڸ�Ʈ2</td>
	  <td>
	  	MD��<input type="text" name="mdname2" value="" size="16" class="input_b"><br>
	  	<textarea name="mdcomment2" cols="60" rows="5"></textarea>
	  </td>
	</tr>
	<tr>
	  <td width="100">�ڸ�Ʈ3</td>
	  <td>
	  	MD��<input type="text" name="mdname3" value="" size="16" class="input_b"><br>
	  	<textarea name="mdcomment3" cols="60" rows="5"></textarea>
	  </td>
	</tr>
// -->
	<tr>
	  <td colspan="2" align="center">
	  	<input type="button" value="����" onClick="SubmitForm();">
	  </td>
	</tr>
	</form>
</table>
<%
end if
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
