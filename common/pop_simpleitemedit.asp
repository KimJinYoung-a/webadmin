<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ��ǰ����
' Hieditor : ������ ����
'			 2021.03.24 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_2008.asp"-->
<%
dim itemid, menupos, i, defaultcheck
	itemid = getNumeric(requestCheckVar(request("itemid"),10))
	menupos = requestCheckVar(request("menupos"),10)

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
Dim vQuery, vIsOK
''vQuery = "EXEC [db_item].[dbo].[sp_Ten_ItemNotificationRaw_Check] '" & itemid & "'"
''rsget.open vQuery,dbget,1
''2015/06/18
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
set oitem = new CItem
	oitem.FRectItemID = itemid
	oitem.FRectSellReserve ="Y"
	'if itemid<>"" then //��ǰ��ȣ �� üũ ����� �Ǵ��� Ȯ������ �ּ�ó�� 2014.03.11 ������
		oitem.GetOneItem
	'end if

If oitem.FTotalCount < 1 Then
	response.write "<script type='text/javascript'>"
	response.write "	alert('�������� �ʴ� ��ǰ�ڵ��Դϴ�');"
	response.write "	self.close();"
	response.write "</script>"
	dbget.close()	:	response.End
End IF

dim oitemoption
set oitemoption = new CItemOption
	oitemoption.FRectItemID = itemid
	if itemid<>"" then
		oitemoption.GetItemOptionInfo
	end if

dim oitemrackoption
set oitemrackoption = new CItemOption
	oitemrackoption.FPageSize = 50
	oitemrackoption.FCurrPage = 1
	oitemrackoption.frectitemgubun = "10"
	oitemrackoption.FRectItemID = itemid

	if itemid<>"" then
		oitemrackoption.GetItemrackcodeInfo
	end if

''�귣�� ���ڵ�
dim sqlStr, prtidx
dim objCmd, returnValue
if (itemid<>"") and (oitem.FResultCount>0) then
    sqlStr = "select prtidx from [db_user].[dbo].tbl_user_c "
    sqlStr = sqlStr & " where userid='" & oitem.FOneItem.FMakerid & "'"
    rsget.Open sqlStr, dbget, 1
    if Not rsget.Eof then
        prtidx = rsget("prtidx")

        prtidx = format00(4,prtidx)
    end if
    rsget.close

'### ���¿��� ���� üũ
if oitem.FOneItem.Fdeliverytype = "1" or  oitem.FOneItem.Fdeliverytype = "4" then '�ٹ����� ����϶� ����� Ȯ��
set objCmd = Server.CreateObject("ADODB.COMMAND")
	With objCmd
		.ActiveConnection = dbget
		.CommandType = adCmdText
		.CommandText = "{?= call db_item.[dbo].[sp_Ten_item_sellreserve_chkStock]("&itemid&")}"
		.Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue)
		.Execute, , adExecuteNoRecords
		End With
	    returnValue = objCmd(0).Value
Set objCmd = nothing
else
	returnValue = 1
end if
end if
'response.write "type="&oitem.FOneItem.Fdeliverytype&"rV="&returnValue
if oitem.FOneItem.Fsellreservedate ="" or isnull(oitem.FOneItem.Fsellreservedate) then
	oitem.FOneItem.Fsellreservedate=now()
	defaultcheck=false
else
	defaultcheck=true
end if

%>
<script language='javascript' src="/js/jsCal/js/jscal2.js"></script>
<script language='javascript' src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />
<script type="text/javascript" src="/js/datetime.js?v=1.0"></script>
<script type="text/javascript">

//���� or ������ Radio��ư Ŭ����
function EnabledCheck(comp){
	var frm = document.frm2;

	for (i = 0; i < frm.elements.length; i++) {
		  var e = frm.elements[i];
		  if ((e.type == 'text') && (e.name.substring(0,"optremainno".length) == "optremainno")) {
				e.disabled = (comp.value=="N");
		  }
  	}

    frm.recalcuLimit.disabled = (comp.value=="N");

    if (comp.value=="N"){ //������
        resetLimit2Zero();
        document.all.dvDisp.style.display = "none";
        frm.limitdispyn[0].checked = false;
        frm.limitdispyn[1].checked = true;
    }else{ //����
        resetLimit();
        document.all.dvDisp.style.display = "";
    }
}

//�������� �缳��
function resetLimit(){
    var frm = document.frm2;

    for (i = 0; i < frm.elements.length; i++) {
		var e = frm.elements[i];
		if ((e.type == 'text')||(e.type == 'radio')) {
		  	if ((e.name.substring(0,"optremainno".length)) == "optremainno"){
		  	    //Enable �� ��츸
		  	    if (!e.disabled){
		  	        //���� ����� 98%�� ���� (��� 10�� �̻��� ��츸) ����(97% -> 98%, 2014-07-25)
		  	        //if (e.getAttribute("dumistock")>=10){
		  	          //  e.value = parseInt(e.getAttribute("dumistock")*0.98);
		  	       // }

					//2016.04.08 ���� ��� ������ �����ϰ� �� ����		'/2016.04.08 ������ �߰�
					//��������� 0���� ������� 0�� ����	'/2016.04.20 �ѿ�� �߰�
					if ( parseInt(e.getAttribute("dumistock"))<0 ){
						e.value = 0;
					}else{
						e.value = parseInt(e.getAttribute("dumistock"));
					}
		  	    }
		  	}
		}
  	}
}

//�������� 0���� Setting
function resetLimit2Zero(){
    var frm = document.frm2;

    for (i = 0; i < frm.elements.length; i++) {
		var e = frm.elements[i];
		if ((e.type == 'text')||(e.type == 'radio')) {
		  	if ((e.name.substring(0,"optremainno".length)) == "optremainno"){
		  	    e.value = 0;
		  	}
		}
  	}
}

function SaveItem(frm){
	var obj, subobj;
    var i, optdanjongyn, optisusing0, optisusing1;

	if ((frm.itemrackcode.value.length > 0) && (frm.itemrackcode.value.length != 4) && (frm.itemrackcode.value.length != 8)){
		alert('��ǰ ���ڵ�� 4�ڸ� �Ǵ� 8�ڸ��� �����Ǿ��ֽ��ϴ�.');
		frm.itemrackcode.focus();
		return;
	}

    if ((frm.subitemrackcode.value.length > 0) && (frm.subitemrackcode.value.length != 4) && (frm.subitemrackcode.value.length != 8)){
		alert('��ǰ �������ڵ�� 4�ڸ� �Ǵ� 8�ڸ��� �����Ǿ��ֽ��ϴ�.');
		frm.subitemrackcode.focus();
		return;
	}

	<% if oitem.FResultCount>0 then %>
	    <% if Not oitem.FOneItem.IsUpchebeasong then %>
	    // �ּ�ó�� - 2014.04.01 ������
	    //�Ǹ� N �ΰ�� ����ǰ�� �Ǵ� MDǰ���� ���� �ؾ���.
	   // if ((frm.sellyn[2].checked)&&!((frm.danjongyn[1].checked)||(frm.danjongyn[2].checked)||(frm.danjongyn[3].checked))){
	     //   alert('�Ǹ� ���� ��ǰ�ΰ�� ������,����ǰ�� �Ǵ� MDǰ���� �����ϼž� �մϴ�.');
	       // frm.danjongyn[2].focus();
	        //return;
	    //}

	    //������,���������� �����Ǹ��ΰ�츸 ������ (�ǸŽø����� ����)
		if ((frm.danjongyn[1].checked)||(frm.danjongyn[2].checked)||(frm.danjongyn[3].checked)){
			if ((frm.sellyn[0].checked)&&(!frm.limityn[0].checked)){
				alert('�Ǹ����̰�, ���� �Ǹ��� ��츸 ������,����ǰ��, MDǰ���� ���� �� �� �ֽ��ϴ�.');
				frm.limityn[0].focus();
				return;
			}
		}
    	<% if oitemoption.FResultCount > 0 then %>
        for (i = 0; ;i++) {
            optisusing0 = document.getElementById('optisusing0_' + i);
            optisusing1 = document.getElementById('optisusing1_' + i);
            optdanjongyn = document.getElementById('optdanjongyn_' + i);

            if (optdanjongyn == undefined) { break; }
            if (optisusing1.checked == true) { continue; }

            if ((optdanjongyn.value == 'S') || (optdanjongyn.value == 'Y') || (optdanjongyn.value == 'M')) {
	    		if ((frm.sellyn[0].checked)&&(!frm.limityn[0].checked)){
		    		alert('�Ǹ����̰�, ���� �Ǹ��� ��츸 ������,����ǰ��, MDǰ���� ���� �� �� �ֽ��ϴ�.');
			    	frm.limityn[0].focus();
				    return;
    			}
            }
        }
    	<% end if %>
    	<% end if %>
	<% end if %>

	//�������̳� �����ϴ°��
	if ((frm.isusing[1].checked)&&(frm.sellyn[0].checked)){
        alert('��� ���� ��ǰ�� �Ǹŷ� ���� �Ұ��մϴ�.');
        frm.sellyn[2].focus();
        return;
    }

	//���¿���
	if(typeof(frm.chkSR)=="object"){
		if(frm.chkSR.checked){
	    if(frm.dSR.value==""){
	    	alert("���¿����� �����Ǿ��ֽ��ϴ�. ��¥�� �Է����ּ���");
	    	frm.dSR.focus();
	    	return;
	    }

		if(toDate(frm.dSR.value+" "+frm.settime.value+":00:00") <= toDate("<%=date() &" "& Num2Str(hour(now()),2,"0","R") & ":00:00"%>")){
		 	alert("��¥/�ð� ������ �߸��Ǿ����ϴ�.\n\n�� ���� ��¥�� �ð��� ���� ���ķ� �������ּ���.");
			frm.dSR.focus();
			return;
		}

	    if(frm.sellyn[0].checked){
		 	 	if(confirm(frm.dSR.value+"�� ���¿���� ��ǰ�Դϴ�. �Ǹ������� ���� �����Ͻø�, ��ǰ���¿��༳���� ��ҵ˴ϴ�. ����Ͻðڽ��ϱ�? ")){
		 	 		frm.dSR.value = "";
		 	 		frm.chkSR.checked= false;
		 	 	}else{
		 	 		frm.sellyn[0].focus();
		 	 		return;
		 	 	}
	 		}

		 	if(frm.sellyn[1].checked){
			 	 	if(confirm(frm.dSR.value+"�� ���¿���� ��ǰ�Դϴ�. �Ͻ�ǰ���� ���� �����Ͻø�, ��ǰ���¿��༳���� ��ҵ˴ϴ�. ����Ͻðڽ��ϱ�? ")){
			 	 		frm.dSR.value = "";
			 	 		frm.chkSR.checked= false;
			 	 	}else{
			 	 		frm.sellyn[1].focus();
			 	 		return;
			 	 	}
		 	}

	 		if(frm.chkSRC.value==0){
	 			alert("�ٹ����� ����� ���, �԰� Ȯ�� �� ���¿����� �����մϴ�.");
	 			frm.chkSR.focus();
	 			return;
	 		}
	   }
	}

	frm.itemoptionarr.value = "";
	//�ɼ� ���� ���� ����
	frm.optremainnoarr.value = "";
	frm.optrackcodearr.value = "";
    frm.suboptrackcodearr.value = "";
	//�ɼ� ��� ����
	frm.optisusingarr.value = "";
    frm.optdanjongynarr.value = '';

    var option_isusing_count = 0;
	var curritemoption;
	for (i = 0; i < frm.elements.length; i++) {
		var e = frm.elements[i];
		if ((e.type == 'text')||(e.type == 'radio')) {
		  	if ((e.name.substring(0,"optremainno".length)) == "optremainno"){
				curritemoption = e.id;
		  	    //���ڸ� ����
		  	    if (!IsDigit(e.value)){
		  	        alert('���� ������ ���ڸ� �����մϴ�.');
		  	        e.select();
		  	        e.focus();
		  	        return;
		  	    }

				frm.itemoptionarr.value = frm.itemoptionarr.value + curritemoption + "," ;
				frm.optremainnoarr.value = frm.optremainnoarr.value + e.value + "," ;

				if (e.id == "0000") {
				    option_isusing_count = 1;
                } else {
					obj = document.getElementById("optrackcode" + curritemoption);
                    subobj = document.getElementById("suboptrackcode" + curritemoption);
					if ((obj.value.length > 0) && (obj.value.length != 4) && (obj.value.length != 8)){
						alert('��ǰ �ɼ� ���ڵ�� 4�ڸ� �Ǵ� 8�ڸ��� �����Ǿ��ֽ��ϴ�.');
						obj.focus();
						return;
					}
                    if ((subobj.value.length > 0) && (subobj.value.length != 4) && (subobj.value.length != 8)){
						alert('��ǰ �ɼ� ���� ���ڵ�� 4�ڸ� �Ǵ� 8�ڸ��� �����Ǿ��ֽ��ϴ�.');
						subobj.focus();
						return;
					}
					frm.optrackcodearr.value = frm.optrackcodearr.value + obj.value + "," ;
                    frm.suboptrackcodearr.value = frm.suboptrackcodearr.value + subobj.value + "," ;
				}
		  	}

            //�ɼ� ��뿩��
			if ((e.name.substring(0,"optisusing".length)) == "optisusing") {
				if (e.checked) {
					if (e.value == "Y") {
					    option_isusing_count = option_isusing_count + 1;
                    }
					frm.optisusingarr.value = frm.optisusingarr.value + e.value + "," ;
				}
			}
		} else if (e.type == 'select-one') {
			if ((e.name.substring(0,"optdanjongyn".length)) == "optdanjongyn") {
				frm.optdanjongynarr.value = frm.optdanjongynarr.value + e.value + "," ;
			}
        }
  	}

    if (option_isusing_count < 1) {
        alert("��� �ɼ��� ���������� �Ҽ� �����ϴ�. ��ǰ������ ���������� �����ϰų�, ���þ��� �����ϼ���.");
        return;
    }

	<%
	If vIsOK = "x" Then
		If oitem.FOneItem.FSellYn <> "Y" Then
	%>
			if(frm.sellyn[0].checked)
			{
				var ret = confirm('��ǰ��ó����� ��� �ԷµǾ� ���� ���� �����Դϴ�.\n�׷��� �Ǹ������� ���� �Ͻðڽ��ϱ�?');
			}
			else
			{
				var ret = confirm('���� �Ͻðڽ��ϱ�?');
			}
	<%	Else %>
			var ret = confirm('���� �Ͻðڽ��ϱ�?');
	<%	End If
	Else
	%>
		var ret = confirm('���� �Ͻðڽ��ϱ�?');
	<% End If %>

	if(ret){
		frm.submit();
	}
}

function popoptionEdit(iid){
	var popwin = window.open('/common/pop_adminitemoptionedit.asp?itemid=' + iid,'popitemoptionedit','width=1200 height=800 scrollbars=yes resizable=yes');
	popwin.focus();
}

function jsPopItemHistory(itemid){
	var popwin = window.open('/common/pop_itemhistory.asp?itemid=' + itemid,'jsPopItemHistory','width=1400 height=800 scrollbars=yes resizable=yes');
	popwin.focus();
}

//�޷�
function jsPopCal(sName){
 if(!document.all.chkSR.checked){
 	 document.all.chkSR.checked= true;
 	}
	var winCal;
	winCal = window.open('/lib/common_cal.asp?DN='+sName,'pCal','width=250, height=200');
	winCal.focus();
}

//���¿���
function jsChkSellReserve(){
	if(!document.all.chkSR.checked){
		document.all.dSR.value = "";
	}
}

function CloseWindow() {
    window.close();
}

function ReloadWindow() {
    document.location.reload();
}

window.resizeTo(1400,800);

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="" style="margin:0px;">
<input type="hidden" name="page" value="1">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		��ǰ�ڵ� : <input type="text" name="itemid" value="<%= itemid %>" Maxlength="9" size="9">
	</td>	
	<td width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="frm.submit();">
	</td>
</tr>
</table>
</form>
<!-- �˻� �� -->

<% if oitem.FResultCount>0 then %>
	<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor=#BABABA>
	<form name=frm2 method=post action="do_simpleiteminfoedit.asp">
	<input type=hidden name=menupos value="<%= menupos %>">
	<input type=hidden name=itemid value="<%= itemid %>">
	<input type=hidden name=itemoptionarr value="">
	<input type=hidden name=optisusingarr value="">
    <input type=hidden name=optdanjongynarr value="">
	<input type=hidden name=optremainnoarr value="">
	<input type=hidden name="optrackcodearr" value="">
    <input type=hidden name="suboptrackcodearr" value="">
	<input type="hidden" name="deliverytype" value="<%=oitem.FOneItem.Fdeliverytype%>">
	<input type="hidden" name="chkSRC" value="<%=returnValue%>">
	<tr>
	<td colspan="2" bgcolor="#FFFFFF">
			<table width="100%" cellspacing=1 cellpadding=1 border="0" class=a bgcolor=#BABABA>
			<tr height="25">

		<td width="120" bgcolor="#DDDDFF">��ǰ��</td>
		<td colspan="2" bgcolor="#FFFFFF"><%= oitem.FOneItem.Fitemname %></td>
	</tr>
	<tr height="25">
		<td bgcolor="#DDDDFF">�귣��ID/�귣���</td>
		<td colspan="2" bgcolor="#FFFFFF">
			<%= oitem.FOneItem.Fmakerid %>/<%= oitem.FOneItem.FBrandName %>
			&nbsp;&nbsp;
			�귣�巢�ڵ� : <%= prtidx %>
		</td>
	</tr>
	<tr height="25">
		<td bgcolor="#DDDDFF">�Һ��ڰ�/���԰�</td>
		<td colspan="2" bgcolor="#FFFFFF">
			<%= FormatNumber(oitem.FOneItem.Forgprice,0) %> / <%= FormatNumber(oitem.FOneItem.Forgsuplycash,0) %>
			&nbsp;&nbsp;
			<font color="<%= mwdivColor(oitem.FOneItem.FMwDiv) %>"><%= oitem.FOneItem.getMwDivName %></font>
			&nbsp;
			<% if oitem.FOneItem.FSellcash<>0 then %>
			<%= CLng((1- oitem.FOneItem.Forgsuplycash/oitem.FOneItem.Forgprice)*100) %> %
			<% end if %>
		</td>
	</tr>

	<% if (oitem.FOneItem.FSailYn="Y") then %>
		<tr height="25">
			<td bgcolor="#DDDDFF">���ΰ�/���԰�</td>
			<td colspan="2" bgcolor="#FFFFFF">
				<font color="red">
					<%= FormatNumber(oitem.FOneItem.FSellcash,0) %> / <%= FormatNumber(oitem.FOneItem.FBuycash,0) %>
					&nbsp;&nbsp;
					<% if (oitem.FOneItem.Forgprice<>0) then %>
				        <%= CLng((oitem.FOneItem.Forgprice-oitem.FOneItem.Fsellcash)/oitem.FOneItem.Forgprice*100) %>%
				    <% end if %>
				    ����
				</font>
				&nbsp;&nbsp;
				<font color="<%= mwdivColor(oitem.FOneItem.FMwDiv) %>"><%= oitem.FOneItem.getMwDivName %></font>
				&nbsp;
				<% if oitem.FOneItem.FSellcash<>0 then %>
					<%= CLng((1- oitem.FOneItem.FBuycash/oitem.FOneItem.FSellcash)*100) %> %
				<% end if %>
			</td>
		</tr>
	<% end if %>

	<% if (oitem.FOneItem.Fitemcouponyn="Y") then %>
		<tr height="25">
			<td bgcolor="#DDDDFF">������/���԰�</td>
			<td colspan="2" bgcolor="#FFFFFF">
				<font color="green">
					<%= FormatNumber(oitem.FOneItem.GetCouponAssignPrice,0) %>
					&nbsp;&nbsp;
					<%= oitem.FOneItem.GetCouponDiscountStr %> ����
				</font>
			</td>
		</tr>
	<% end if %>

	<tr height="25">
		<td bgcolor="#DDDDFF">��ǰ���ڵ�</td>
		<td bgcolor="#FFFFFF" width="270">
			<input type="text" name="itemrackcode" value="<%= oitem.FOneItem.FitemRackCode %>" size="8" maxlength="8" > (4 or 8�ڸ� Fix)
		</td>
		<td rowspan="5" align="right" bgcolor="#FFFFFF">
			<img src="<%= oitem.FOneItem.FListImage %>" width="100" align="right">
		</td>
	</tr>
	<tr height="25">
		<td bgcolor="#DDDDFF">�������ڵ�</td>
		<td bgcolor="#FFFFFF">
		    <input type="text" name="subitemrackcode" value="<%= oitem.FOneItem.Fsubitemrackcode %>" size="8" maxlength="8" > (4 or 8�ڸ� Fix)
		</td>
	</tr>
	<tr height="25">
		<td bgcolor="#DDDDFF">���ɼ�</td>
		<td bgcolor="#FFFFFF">
		(<%= oitem.FOneItem.FOptionCnt %> ��)
		&nbsp;
		<input type=button class="button" value="�ɼ��߰�/����" onclick="popoptionEdit('<%= itemid %>');">
		</td>
	</tr>
	<tr height="25">
		<td bgcolor="#DDDDFF">��۱���</td>
		<td bgcolor="#FFFFFF">
		<% if oitem.FOneItem.IsUpcheBeasong then %>
		<b>��ü</b>���
		<% else %>
		�ٹ����ٹ��
		<% end if %>
		</td>
	</tr>
	<tr height="25">
		<td bgcolor="#DDDDFF">��ǰ ǰ������</td>
		<td bgcolor="#FFFFFF">
		<% if oitem.FOneItem.IsSoldOut then %>
		<font color=red><b>ǰ��</b></font>
		<% end if %>
		</td>
	</tr>
	<tr height="25">
		<td bgcolor="#DDDDFF">��� ��ۼҿ���</td>
		<td bgcolor="#FFFFFF" colspan="2">

			<% if (oitem.FOneItem.FavgDLvDate>-1) then %>
			    D+<%= oitem.FOneItem.FavgDLvDate+1 %>
			<% else %>
			    ������ ����
			<% end if %>
			&nbsp;&nbsp;&nbsp;
			<a href="javascript:popItemAvgDlvGraph('<%= itemid %>');">[�����׷���]</a>&nbsp;
			<a href="javascript:popItemAvgDlvList('<%= itemid %>');">[�󼼸���Ʈ]</a>
		</td>
	</tr>
	<tr height="25">
		<td bgcolor="#DDDDFF">�Ǹ� ������</td>
		<td bgcolor="#FFFFFF" colspan="2">
			<%= oitem.FOneItem.FsellSTDate %>
		</td>
	</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td colspan="2" bgcolor="#FFFFFF">
			<table width="100%" cellspacing=1 cellpadding=1 class=a bgcolor=#BABABA>
			<tr height="25">
				<td width="120" bgcolor="#DDDDFF">��ǰ �Ǹſ���</td>
				<td bgcolor="#FFFFFF">
					<% if oitem.FOneItem.FSellYn="Y" then %>
					<input type="radio" name="sellyn" value="Y" checked >�Ǹ���
					<input type="radio" name="sellyn" value="S" >�Ͻ�ǰ��
					<input type="radio" name="sellyn" value="N" >�Ǹž���
					<% elseif oitem.FOneItem.FSellYn="S" then %>
					<input type="radio" name="sellyn" value="Y" >�Ǹ���
					<input type="radio" name="sellyn" value="S" checked ><font color="red">�Ͻ�ǰ��</font>
					<input type="radio" name="sellyn" value="N" >�Ǹž���
					<% else %>
					<input type="radio" name="sellyn" value="Y" >�Ǹ���
					<input type="radio" name="sellyn" value="S" >�Ͻ�ǰ��
					<input type="radio" name="sellyn" value="N" checked ><font color="red">�Ǹž���</font>
					<% end if %>

                    <input type="button" class="button" value="�����丮" onClick="jsPopItemHistory(<%= itemid %>)">
				</td>
			</tr>
			<tr height="25">
				<td bgcolor="#DDDDFF">��ǰ ��뿩��</td>
				<td bgcolor="#FFFFFF">
					<% if oitem.FOneItem.FIsUsing="Y" then %>
					<input type="radio" name="isusing" value="Y" checked >�����
					<input type="radio" name="isusing" value="N" >������
					<% else %>
					<input type="radio" name="isusing" value="Y" >�����
					<input type="radio" name="isusing" value="N" checked ><font color="red">������</font>
					<% end if %>
				</td>
			</tr>
			<tr height="25">
				<td bgcolor="#DDDDFF">���� ��뿩��</td>
				<td bgcolor="#FFFFFF">
					<% if oitem.FOneItem.FIsExtUsing="Y" then %>
					<input type="radio" name="isExtusing" value="Y" checked >�����
					<input type="radio" name="isExtusing" value="N" >������
					<% else %>
					<input type="radio" name="isExtusing" value="Y" >�����
					<input type="radio" name="isExtusing" value="N" checked ><font color="red">������</font>
					<% end if %>
				</td>
			</tr>
			<tr height="25">
				<td bgcolor="#DDDDFF">��ǰ ��������</td>
				<td bgcolor="#FFFFFF">
					<% if oitem.FOneItem.Fdanjongyn="Y" then %>
						<input type="radio" name="danjongyn" value="N" >������
						<input type="radio" name="danjongyn" value="S" >������
						<input type="radio" name="danjongyn" value="Y" checked ><font color="red">����ǰ��</font>
						<input type="radio" name="danjongyn" value="M" >MDǰ��
					<% elseif oitem.FOneItem.Fdanjongyn="S" then %>
						<input type="radio" name="danjongyn" value="N" >������
						<input type="radio" name="danjongyn" value="S" checked ><font color="red">������</font>
						<input type="radio" name="danjongyn" value="Y" >����ǰ��
						<input type="radio" name="danjongyn" value="M" >MDǰ��
					<% elseif oitem.FOneItem.Fdanjongyn="M" then %>
						<input type="radio" name="danjongyn" value="N" >������
						<input type="radio" name="danjongyn" value="S" >������
						<input type="radio" name="danjongyn" value="Y" >����ǰ��
						<input type="radio" name="danjongyn" value="M" checked ><font color="red">MDǰ��</font>
					<% else %>
						<input type="radio" name="danjongyn" value="N" checked >������
						<input type="radio" name="danjongyn" value="S" >������
						<input type="radio" name="danjongyn" value="Y" >����ǰ��
						<input type="radio" name="danjongyn" value="M" >MDǰ��
					<% end if %>
					<font color="#AAAAAA">
					<br> (��ǰ�Ǹſ��� ������� - �߰� �԰��� ������ ��������)
				</font>
				</td>
			</tr>
			<tr height="25">
				<td bgcolor="#DDDDFF">�����Ǹſ���</td>
				<td bgcolor="#FFFFFF">
				<% if oitem.FOneItem.FLimitYn="Y" then %>
				<input type="radio" name="limityn" value="Y" checked onclick="EnabledCheck(this)"><font color="blue">�����Ǹ�</font>
				<input type="radio" name="limityn" value="N" onclick="EnabledCheck(this)">�������Ǹ�
				(<%= oitem.FOneItem.FLimitNo %>-<%= oitem.FOneItem.FLimitSold %>=<%= oitem.FOneItem.FLimitNo-oitem.FOneItem.FLimitSold %>)
				<% else %>
				<input type="radio" name="limityn" value="Y" onclick="EnabledCheck(this)">�����Ǹ�
				<input type="radio" name="limityn" value="N" checked onclick="EnabledCheck(this)">�������Ǹ�
				<% end if %>
				<div id="dvDisp" style="display:<% if oitem.FOneItem.FLimitYn<>"Y" then %>none<%END IF%>;" >&nbsp;-> �������⿩��:
					<input type="radio" name="limitdispyn" value="Y" <%IF oitem.FOneItem.Flimitdispyn="Y" or oitem.FOneItem.Flimitdispyn ="" THEN%>checked<%END IF%>>����
					<input type="radio" name="limitdispyn" value="N" <%IF oitem.FOneItem.Flimitdispyn="N" THEN%>checked<%END IF%>>�����</div>
				</td>
			</tr>
			</table>

		</td>
	</tr>
	<tr>
	    <td colspan="2" bgcolor="#FFFFFF">��������� 10�̸��� ���� ����ľ� �� ����� �Է��Ͻñ� �ٶ��ϴ�.</td>
	</tr>
	<tr>
		<td colspan="2" bgcolor="#FFFFFF">
			<table width="100%" cellspacing=1 cellpadding=1 class=a bgcolor=#BABABA>
			<tr height="25" align="center" bgcolor="#FFDDDD" >
				<td width="50">�ɼ�<br />�ڵ�</td>
				<td>�ɼǸ�</td>
				<td width="70">�ɼ�<br />���<br />����</td>
                <td width="70">�ɼ�<br />����<br />����</td>
				<td width="40">����<br>����</td>
				<td width="80">�����Ǹż���<br><input name="recalcuLimit" type="button" class="button" value="��������" onclick="resetLimit();" <%= chkIIF(oitem.FOneItem.FLimitYn="N","disabled","") %>></td>
				<td width="40"><a href="javascript:TnPopItemStock('<%= itemid %>','');">����<br />��<br />���</a></td>
				<td width="180">����ȣ / ������</td>
			</tr>
			<% if oitemoption.FResultCount>0 then %>
				<% for i=0 to oitemoption.FResultCount - 1 %>
					<% if oitemoption.FITemList(i).FOptIsUsing="N" then %>
					<tr align="center" bgcolor="#EEEEEE">
					<% else %>
					<tr align="center" bgcolor="#FFFFFF">
					<% end if %>
						<td><%= oitemoption.FITemList(i).FItemOption %></td>
						<td><%= oitemoption.FITemList(i).FOptionName %></td>
						<td>
							<% if oitemoption.FITemList(i).FOptIsUsing="Y" then %>
							<input type="radio" id="optisusing0_<%= i %>" name="optisusing<%= oitemoption.FITemList(i).FItemOption %>" value="Y" checked >Y <input type="radio" id="optisusing1_<%= i %>" name="optisusing<%= oitemoption.FITemList(i).FItemOption %>" value="N" >N
							<% else %>
							<input type="radio" id="optisusing0_<%= i %>" name="optisusing<%= oitemoption.FITemList(i).FItemOption %>" value="Y" >Y <input type="radio" id="optisusing1_<%= i %>" name="optisusing<%= oitemoption.FITemList(i).FItemOption %>" value="N" checked ><font color="red">N</font>
							<% end if %>
						</td>
                        <td>
                            <select class="select" id="optdanjongyn_<%= i %>" name="optdanjongyn<%= oitemoption.FITemList(i).FItemOption %>" <%= CHKIIF(oitemoption.FITemList(i).Foptdanjongyn<>"N", "style='background-color:#FFBBDD'", "") %>>
                                <option value="N" <%= CHKIIF(oitemoption.FITemList(i).Foptdanjongyn="N", "selected", "") %>>������</option>
                                <option value="S" <%= CHKIIF(oitemoption.FITemList(i).Foptdanjongyn="S", "selected", "") %>>������</option>
                                <option value="Y" <%= CHKIIF(oitemoption.FITemList(i).Foptdanjongyn="Y", "selected", "") %>>�ɼǴ���</option>
                                <option value="M" <%= CHKIIF(oitemoption.FITemList(i).Foptdanjongyn="M", "selected", "") %>>MDǰ��</option>
                            </select>
                        </td>
						<td><%= oitemoption.FITemList(i).GetOptLimitEa %></td>
						<td>
							<input type="text" id="<%= oitemoption.FITemList(i).FItemOption %>" dumistock="<%= oitemoption.FITemList(i).GetLimitStockNo %>" name="optremainno<%= oitemoption.FITemList(i).FItemOption %>" value="<%= oitemoption.FITemList(i).GetOptLimitEa %>" size="4" maxlength=5 <% if oitem.FOneItem.FLimitYn="N" then response.write "disabled" %> >
						</td>
						<td <%= chkIIF(oitemoption.FITemList(i).GetLimitStockNo<10,"bgcolor='#6666EE'","") %> ><a href="javascript:TnPopItemStock('<%= itemid %>','<%= oitemoption.FITemList(i).FItemOption %>');"><%= oitemoption.FITemList(i).GetLimitStockNo %></a></td>
						<td>
							<input type="text" id="optrackcode<%= oitemoption.FITemList(i).FItemOption %>" name="optrackcode" value="<%= oitemoption.FITemList(i).Foptrackcode %>" size="8" maxlength="8" >
                            <input type="text" id="suboptrackcode<%= oitemoption.FITemList(i).FItemOption %>" name="suboptrackcode" value="<%= oitemoption.FITemList(i).Fsuboptrackcode %>" size="8" maxlength="8" >
						</td>
					</tr>
				<% next %>
			<% else %>
					<tr align="center" bgcolor="#FFFFFF">
						<td>0000</td>
						<td colspan="3">�ɼǾ���</td>
						<td><%= oitem.FOneItem.GetLimitEa %></td>
						<td>
							<input type="text" id="0000" dumistock="<%= oitem.FOneItem.GetLimitStockNo %>" name="optremainno" value="<%= oitem.FOneItem.GetLimitEa %>" size="4" maxlength=5 <% if oitem.FOneItem.FLimitYn="N" then response.write "disabled" %> >
						</td>
						<td <%= chkIIF(oitem.FOneItem.GetLimitStockNo<10,"bgcolor='#6666EE'","") %> >
							<a href="javascript:TnPopItemStock('<%= itemid %>','');"><%= oitem.FOneItem.GetLimitStockNo %></a>
						</td>
						<td></td>
					</tr>
			<% end if %>
			</table>
		</td>
	</tr>

	<% IF oitem.FOneItem.Fsellyn = "N" THEN '�Ǹž��� �����϶��� �����ش� %>
		<tr>
			<td   bgcolor="#FFFFFF">
				<table width="100%" border="0" align="center" class="a" cellpadding="5" cellspacing="0">
			 	<tr>
					<td>
						<input type="checkbox" name="chkSR" value="Y" onClick="jsChkSellReserve();" <%IF defaultcheck THEN%>checked<%END IF%>> ��ǰ���¿���:
						<input type="text" id="dSR" name="dSR" value="<%= FormatDateTime(oitem.FOneItem.Fsellreservedate,2) %>" size="12" class="input" />
						<select name="settime">
							<% for i=0 to 23 %>
							<option value="<%=Format00(2,i)%>"<% if Hour(oitem.FOneItem.Fsellreservedate)=i then response.write " selected" %>><%=Format00(2,i)%></option>
							<% next %>
						</select>��
						<img id="dSR_trigger" src="/images/admin_calendar.png" />
						  <div style="padding:3px">������ ������ ��� ����� �ð��� ������ ���� �ʽ��ϴ�. <br>
					   �ٹ����� ����� ���, �԰� Ȯ�� �� ���¿����� �����մϴ�.  </div>
					   <script type="text/javascript">
						var CAL_SR = new Calendar({
							inputField : "dSR", trigger    : "dSR_trigger",
							onSelect: function() {
								this.hide();
							}
							, min: "<%=date()%>"
							, bottomBar: true, dateFormat: "%Y-%m-%d"
						});
					   </script>
					</td>
					</tR>
				</table>
			</td>
		</tr>
	<% END IF %>

	<tr align="center">
	    <td colspan="2" bgcolor="#FFFFFF">
			<input type="button" value="�����ϱ�" onclick="SaveItem(frm2)" class="button">
			<input type="button" value=" �� �� " onclick="CloseWindow()" class="button">
		</td>
	</tr>
	<input type=hidden name="pojangok" value="<%= oitem.FOneItem.FPojangOK %>">
	</form>
	</table>
<% else %>
	<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor=#BABABA>
	<tr bgcolor="#FFFFFF">
	    <td align="center">[�˻� ����� �����ϴ�.]</td>
	</tr>
	</table>
<% end if %>

<Br>
<% if oitemrackoption.FTotalCount >0 then %>
<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<% if oitemrackoption.FresultCount>0 then %>
<tr height="25" bgcolor="FFFFFF">
	<td colspan="15">
		�� ���ڵ� ���� �α� / �˻���� : <b><%= oitemrackoption.FTotalCount %></b>&nbsp;&nbsp; �� �ִ� 50������ ����˴ϴ�.
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width=80>����</td>
	<td>���泻��</td>
	<td width=70>��ǰ�ڵ�</td>
	<td width=60>�ɼ��ڵ�</td>
	<td>�ɼǸ�</td>
	<td width=80>����<br>���ڵ�</td>
	<td width=80>����<br>�������ڵ�</td>
</tr>
<% for i=0 to oitemrackoption.FresultCount-1 %>
<tr align="center" bgcolor="#FFFFFF">
	<td>
		<%= left(oitemrackoption.FItemList(i).fregdate,10) %>
		<br><%= mid(oitemrackoption.FItemList(i).fregdate,12,22) %>
		<br><%= oitemrackoption.FItemList(i).fregadminid %>
	</td>
	<td align="left">
		<%= oitemrackoption.FItemList(i).fcomment %>
	</td>
	<td>
		<%= oitemrackoption.FItemList(i).fitemid %>
	</td>
	<td>
		<%= oitemrackoption.FItemList(i).fitemoption %>
	</td>
	<td align="left">
		<%= oitemrackoption.FItemList(i).Fitemoptionname %>
	</td>
	<td>
		<%= oitemrackoption.FItemList(i).frackcodeByOption %>
	</td>
	<td>
		<%= oitemrackoption.FItemList(i).fsubRackcodeByOption %>
	</td>
</tr>   
<% next %>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="15" align="center" class="page_link">[�˻������ �����ϴ�.]</td>
	</tr>
<% end if %>
</table>
<% end if %>

<%
set oitemoption = Nothing
set oitem = Nothing
set oitemrackoption = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
