<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ��ǰ����
' Hieditor : 2009.04.07 ������ ����
'			 2011.04.28 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/partner/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" --> 
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/new_itemcls.asp"-->
<!-- #include virtual="/partner/lib/adminHead.asp" --><!--html--> 
<%
dim itemid ,i
	itemid = requestCheckvar(request("itemid"),16)  ''requestCheckvar 2016/02/11

if itemid = "" then
	response.write "<script>"
	response.write "	alert('��ǰ�ڵ尡 �����ϴ�');"
	response.write "	self.close();"
	response.write "</script>"
	dbget.close()	:	response.End
end if


'####### ��ǰ��ù��� ���� ������üũ
If IsNumeric(itemid) = false Then
	response.write "<script>"
	response.write "	alert('�߸��� ��ǰ�ڵ��Դϴ�');"
	response.write "	self.close();"
	response.write "</script>"
	dbget.close()	:	response.End
End IF

''2017/06/19 �߰� by eastone  itemid=7171719721 ����?
If LEN(itemid)>9 Then
	response.write "<script>"
	response.write "	alert('�߸��� ��ǰ�ڵ��Դϴ�');"
	response.write "	self.close();"
	response.write "</script>"
	dbget.close()	:	response.End
End IF

Dim vQuery, vIsOK
''vQuery = "EXEC [db_item].[dbo].[sp_Ten_ItemNotificationRaw_Check] '" & itemid & "'"
''rsget.open vQuery,dbget,1
''2015/06/18 ������
vQuery = "[db_item].[dbo].[sp_Ten_ItemNotificationRaw_Check]('" & itemid & "')"
rsget.CursorLocation = adUseClient
rsget.Open vQuery, dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
If Not rsget.Eof Then
	vIsOK = rsget(0)
Else
	vIsOK = "x"
End IF
rsget.close()
'rw vIsOK
'####### ��ǰ��ù��� ���� ������üũ


dim oitem
set oitem = new CItemInfo

oitem.FRectItemID = itemid

if itemid<>"" then
	oitem.GetOneItemInfo
end if

''2016/02/11 �߰�.
if (oitem.FResultCount<1) then
    response.write "<script>"
	response.write "	alert('�߸��� ��ǰ�ڵ��̰ų� �ش��ǰ�� �����ϴ�.');"
	response.write "	self.close();"
	response.write "</script>"
	dbget.close()	:	response.End
end if

dim oitemoption
set oitemoption = new CItemOption
oitemoption.FRectItemID = itemid
if itemid<>"" then
	oitemoption.GetItemOptionInfo
end if
%>

<script language='javascript'>

function EnabledCheck(comp){
	var frm = document.frm2;

	for (i = 0; i < frm.elements.length; i++) {
		  var e = frm.elements[i];
		  if ((e.type == 'text') && ((e.name.substring(0,"optlimitno".length) == "optlimitno")||(e.name.substring(0,"optlimitsold".length) == "optlimitsold"))) {
				e.disabled = (comp.value=="N");
		  }
  	}

}

function SaveItem(frm){
	frm.itemoptionarr.value = ""
	frm.optlimitnoarr.value = ""
	frm.optlimitsoldarr.value = ""
	frm.optisusingarr.value = ""

    var option_isusing_count = 0;
	for (i = 0; i < frm.elements.length; i++) {
		var e = frm.elements[i];
		if ((e.type == 'text')||(e.type == 'radio')) {
		  	if ((e.name.substring(0,"optlimitno".length)) == "optlimitno"){

		  	    if (!IsDigit(e.value)){
		  	        alert('���������� ���ڸ� �����մϴ�.');
		  	        e.focus();
		  	        return;
		  	    }

				frm.itemoptionarr.value = frm.itemoptionarr.value + e.id + "," ;
				frm.optlimitnoarr.value = frm.optlimitnoarr.value + e.value + "," ;

				if (e.id == "0000") {
				    option_isusing_count = 1;
                }
		  	}

		  	if ((e.name.substring(0,"optlimitsold".length)) == "optlimitsold") {
		  	    if (!IsDigit(e.value)){
		  	        alert('���������� ���ڸ� �����մϴ�.');
		  	        e.focus();
		  	        return;
		  	    }

				frm.optlimitsoldarr.value = frm.optlimitsoldarr.value + e.value + "," ;
			}

			if ((e.name.substring(0,"optisusing".length)) == "optisusing") {
				if (e.checked) {
					if (e.value == "Y") {
					    option_isusing_count = option_isusing_count + 1;
                    }
					frm.optisusingarr.value = frm.optisusingarr.value + e.value + "," ;
				}
			}
		}
  	}
    if (option_isusing_count < 1) {
        alert("��� �ɼ��� ���������� �Ҽ� �����ϴ�. ��ǰ �Ǹſ��θ� �Ǹž������� �������ּ���");
        //alert(frm.itemoptionarr.value);
        return;
    }

<% if (oitem.FOneItem.Fmwdiv <> "U") then %>
    if (frm.reqstring.value == "") {
        alert("���������� �Է����ּ���.");
        return;
    }
<% end if %>

<%
	If vIsOK = "x" Then
		If oitem.FOneItem.FSellYn <> "Y" Then
%>
			if(frm.sellyn[0].checked)
			{
				alert("��ǰ��ó����� ��� �ԷµǾ� ���� ���� �����Դϴ�.\n��� �Է��ϼž� �Ǹ������� ���� �����մϴ�.\n��� �Է��Ͻ� �� �� â�� ���� ���ðų� ���ΰ�ħ �Ͻø� ������ �����մϴ�.");
				return;
			}
<%
		End If
	End If
%>

	var ret = confirm('���� �Ͻðڽ��ϱ�?');

	if(ret){
		frm.submit();
	}
}

function PopOptionEdit(itemid){
	var popwin = window.open('/common/partner/pop_adminitemoptionedit.asp?itemid=' + itemid,'PopOptionEdit','width=800 height=500 scrollbars=yes resizable=yes');
	popwin.focus();
}

function CloseWindow() {
    window.close();
}

function editItemInfo(itemid) {
	var param = "itemid=" + itemid;

	popwin = window.open('/partner/itemmaster/item_infomodify.asp?' + param ,'editItemInfo','width=1100,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}
</script>
</head>
<body>
<div class="popupWrap">
	<div class="popHead">
		<h1><img src="/images/partner/pop_admin_logo.gif" alt="10x10" /></h1>
		<p class="btnClose"><input type="image" src="/images/partner/pop_admin_btn_close.gif" alt="â�ݱ�" onclick="window.close();" /></p>
	</div>
	<div class="popContent scrl">
		<div class="contTit bgNone"><!-- for dev msg : Ÿ��Ʋ �����ϴܿ� searchWrap�� �� ��쿣 bgNone Ŭ���� ���� -->
			<h2>��ǰ����</h2>
			<ul class="txtList">
				<li>�ٹ�(�ٹ����ٹ��)��ǰ�� �ٹ����� Ȯ���� <span class="cRd1">������ ���������� �ݿ�</span>�˴ϴ�.</li>
				<li>����(��ü���) ��ǰ�� ��� <span class="cRd1">��ùݿ�</span>�˴ϴ�.</li>
				<li>�����̳�, ��ǰ�� �� ��Ÿ ���� �Ͻ� ������ <span class="cRd1">��翥��</span>���� ������ �ּ���.</li>
			</ul>  
			</div>
		<div class="cont">  
			<h3>�ɼ�/����/�ǸŰ���</h3>
			<% if oitem.FResultCount>0 then %> 
			<form name="frm2" method="post" action="/common/partner/do_upche_simpleiteminfoedit.asp">
			<input type="hidden" name="itemid" value="<%= itemid %>">
			<input type="hidden" name="itemoptionarr" value="">
			<input type="hidden" name="optisusingarr" value="">
			<input type="hidden" name="optlimitnoarr" value="">
			<input type="hidden" name="optlimitsoldarr" value="">
			<table class="tbType1 writeTb tMar10">
					<colgroup>
						<col width="15%" /><col width="" />
					</colgroup>
					<tbody> 
					<tr> 
						<th><div>��ǰ�ڵ�</div></td>
						<td><%= itemid %></td>
					</tr>
					<tr>
						<th><div>��ǰ��</div></td>
						<td><%= oitem.FOneItem.Fitemname %></td>	
					</tr>
					<tr>
						<th><div>�귣��</div></td>
						<td><%= oitem.FOneItem.Fmakerid %> (<%= oitem.FOneItem.FBrandName %>)</td>
					</tr>
					<tr>
						<th><div>�ǸŰ�/���԰�</div></td>
						<td>
						<%= FormatNumber(oitem.FOneItem.FSellcash,0) %> / <%= FormatNumber(oitem.FOneItem.FBuycash,0) %>
						</td>
					</tr>
					<tr>
						<th><div>���Ա���</div></td>
						<td>
						<font color="<%= oitem.FOneItem.getMwDivColor %>"><%= oitem.FOneItem.getMwDivName %></font>
						&nbsp;
						<% if oitem.FOneItem.FSellcash<>0 then %>
						<%= CLng((1- oitem.FOneItem.FBuycash/oitem.FOneItem.FSellcash)*100) %> %
						<% end if %>
						</td>
					</tr>
					<tr>
						<th><div>���ɼ�</div></td>
						<td>
						(<%= oitem.FOneItem.FOptionCnt %> ��)
						&nbsp;
						<% if oitem.FOneItem.IsUpcheBeasong then %>
						<input type=button value="�ɼǼ���" onclick="PopOptionEdit('<%= itemid %>');" class="btn3 btnIntb">
						<% else %>
						<span class="cRd1">- �ɼ� �߰�/������ ���MD</span>���� �����ϼ���.
						<% end if %>
						</td>
					</tr>
					<tr>
						<th><div>��۱���</div></td>
						<td>
						<% if oitem.FOneItem.IsUpcheBeasong then %>
						<b>��ü</b>���
						<% else %>
						�ٹ����ٹ��
						<% end if %>
						</td>
					</tr>
					<tr>
						<th><div>��ǰ ǰ������</div></td>

						<td>
						<% if (oitem.FOneItem.IsSoldOut) or (oitem.FOneItem.FSellYn="S") then %>
						<span class="cRd1"><strong>ǰ��</strong></span>
						<% end if %>
						</td>
					</tr>
					<tr>
						<th><div>��ǰ �Ǹſ���</div></td>
						<td>
						<% if oitem.FOneItem.FSellYn="Y" then %>
						<input type="radio" name="sellyn" value="Y" class="formRadio" checked >�Ǹ���
						<input type="radio" name="sellyn" value="S" class="formRadio" >�Ͻ�ǰ��
						<input type="radio" name="sellyn" value="N" class="formRadio" >�Ǹž���
						<% elseif oitem.FOneItem.FSellYn="S" then %>
						<input type="radio" name="sellyn" value="Y" class="formRadio" >�Ǹ���
						<input type="radio" name="sellyn" value="S" class="formRadio" checked ><font color="blue">�Ͻ�ǰ��</font>
						<input type="radio" name="sellyn" value="N" class="formRadio" >�Ǹž���
						<% else %>
						<input type="radio" name="sellyn" value="Y" class="formRadio" >�Ǹ���
						<input type="radio" name="sellyn" value="S" class="formRadio" >�Ͻ�ǰ��
						<input type="radio" name="sellyn" value="N" class="formRadio" checked ><font color="red">�Ǹž���</font>
						<% end if %>
						<% If vIsOK = "x" Then %>
					    	&nbsp;&nbsp;<input type="button" class=btn3 btnIntb" value="��ǰ��ó����Է�" style="width:110px;" onClick="editItemInfo('<%=itemid%>');">
						<% End If %>
						</td>
					</tr>
					<input type="hidden" name="isusing" value="<%= oitem.FOneItem.FIsUsing %>">
					<tr>
						<th><div>�����Ǹſ���</div></td>
						<td>
						<% if oitem.FOneItem.FLimitYn="Y" then %>
						<input type="radio" class="formRadio" name="limityn" value="Y" checked onclick="EnabledCheck(this)"><font color="blue">�����Ǹ�</font>
						<input type="radio" class="formRadio" name="limityn" value="N" onclick="EnabledCheck(this)">�������Ǹ�
						<% else %>
						<input type="radio" class="formRadio" name="limityn" value="Y" onclick="EnabledCheck(this)">�����Ǹ�
						<input type="radio" class="formRadio" name="limityn" value="N" checked onclick="EnabledCheck(this)">�������Ǹ�
						<% end if %>
						</td>
					</tr>
				</table>
				<table class="tbType1 listTb">
					<thead>
						<tr> 
								<th><div>�ɼǸ�</div></th>
								<th><div>�ɼǻ�뿩��</div></th>
								<th><div>�������� - �Ǹż��� = �������</div></th>
								<th><div>���</div></th>
						</tr>
					</thead>
					<tbody>
							<% if oitemoption.FResultCount>0 then %>
								<% for i=0 to oitemoption.FResultCount - 1 %>
									<% if oitemoption.FITemList(i).FOptIsUsing="N" then %>
									<tr bgcolor="#EEEEEE">
									<% else %>
									<tr bgcolor="#FFFFFF">
									<% end if %>
										<td><%= oitemoption.FITemList(i).FOptionName %>(<%= oitemoption.FITemList(i).FItemOption %>)</td>
										<td>
											<% if oitemoption.FITemList(i).Foptisusing="Y" then %>
											<input type="radio" class="formRadio" name="optisusing<%= oitemoption.FITemList(i).FItemOption %>" value="Y" checked >����� <input type="radio" class="formRadio" name="optisusing<%= oitemoption.FITemList(i).FItemOption %>" value="N" >������
											<% else %>
											<input type="radio" class="formRadio" name="optisusing<%= oitemoption.FITemList(i).FItemOption %>" value="Y" >����� <input type="radio" class="formRadio" name="optisusing<%= oitemoption.FITemList(i).FItemOption %>" value="N" checked ><span class="cRd1">������</span>
											<% end if %>
										</td>
										<td>
										<input type="text" class="formTxt" id="<%= oitemoption.FITemList(i).FItemOption %>" name="optlimitno<%= oitemoption.FITemList(i).FItemOption %>" value="<%= oitemoption.FITemList(i).FOptLimitNo %>" size="4" maxlength=5 <% if oitem.FOneItem.FLimitYn="N" then response.write "disabled" %> >
										-
										<input type="text" class="formTxt" id="<%= oitemoption.FITemList(i).FItemOption %>" name="optlimitsold<%= oitemoption.FITemList(i).FItemOption %>" value="<%= oitemoption.FITemList(i).FOptLimitSold %>" size="4" maxlength=5 <% if oitem.FOneItem.FLimitYn="N" then response.write "disabled" %> >
										=
										<input type="text" class="formTxt" name="optremainno<%= oitemoption.FITemList(i).FItemOption %>" value="<%= oitemoption.FITemList(i).GetOptLimitEa %>" size="4" maxlength=5 disabled >
									</td>
									<td>
									<% if (oitemoption.FITemList(i).FOptIsUsing="N") or (oitemoption.FITemList(i).Foptsellyn="N") or (oitemoption.FITemList(i).Foptlimityn="Y" and oitemoption.FITemList(i).GetOptLimitEa<1) then %>
									<span class="cRd1">ǰ��</span>
									<% end if %>
									</td>
									</tr>
								<% next %>
							<% else %>
								<tr>
									<td colspan="2">�ɼǾ��� (0000)</td>
									<td>
									<input type="text" class="formTxt" id="0000" name="optlimitno" value="<%= oitem.FOneItem.FLimitNo %>" size="4" maxlength=5 <% if oitem.FOneItem.FLimitYn="N" then response.write "disabled" %> >
									-
									<input type="text" class="formTxt" id="0000" name="optlimitsold" value="<%= oitem.FOneItem.FLimitSold %>" size="4" maxlength=5 <% if oitem.FOneItem.FLimitYn="N" then response.write "disabled" %> >
									=
									<input type="text" class="formTxt" name="optremainno" value="<%= oitem.FOneItem.GetLimitEa %>" size="4" maxlength=5 disabled >
								</td>
								<td>
								    <% if oitem.FOneItem.isSoldOut() then %>
								    <span class="cRd1">ǰ��</span>
								    <% end if %>
								</td>
								</tr>
							<% end if %>
						</tbody>
					</table>
			 <input type="hidden" name="pojangok" value="<%= oitem.FOneItem.FPojangOK %>">
			 <table class="tbType1 writeTb tMar10">
					<colgroup>
						<col width="15%" /><col width="" />
					</colgroup>
					<tbody> 
					<tr> 
						<th><div>�̹���</div></th>
						<td>
							<img src="<%= oitem.FOneItem.FListImage %>" width=100>
						</td>
					</tr>
			<% if (oitem.FOneItem.Fmwdiv <> "U") then %>
			<tr>
				<th><div>��������</div></th>
				<td>
				  <input type="text" class="formTxt" name="reqstring" value="" size="30">
				  <p class="tMar05 fs11">(ex: ����, ����Ͻú���(�԰����� 2003-05-15), ���԰�..)</p>
				</td>
			</tr>
			<% end if %>
			</form>
		</table>
		<div class="tPad15 ct"> 
				<% if (oitem.FOneItem.Fmwdiv = "U") then %>
					<input type="button" value=" �����ϱ� " onclick="SaveItem(frm2)" class="btn3 btnRd" />
				<% else %>
					<input type="button" value=" ������û " onclick="SaveItem(frm2)" class="btn3 btnRd" />
				<% end if %>	
			</div>   
<% end if %>
		</div>
	</div>
</div>
</body>
</html>
<%
set oitemoption = Nothing
set oitem = Nothing
%>
 
<!-- #include virtual="/lib/db/dbclose.asp" -->