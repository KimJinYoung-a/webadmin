<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->

<!-- #include virtual="/lib/classes/new_itemcls.asp"-->
<%
dim itemid
itemid = request("itemid")

dim oitem
set oitem = new CItemInfo
oitem.FRectItemID = itemid

if itemid<>"" then
	oitem.GetOneItemInfo
end if

dim oitemoption
set oitemoption = new CItemOption
oitemoption.FRectItemID = itemid
if itemid<>"" then
	oitemoption.GetItemOptionInfo
end if

dim i

''�귣�� ���ڵ�
dim sqlStr, prtidx
if (itemid<>"") then
    sqlStr = "select prtidx from [db_user].[10x10].tbl_user_c "
    sqlStr = sqlStr & " where userid='" & oitem.FOneItem.FMakerid & "'"
    rsget.Open sqlStr, dbget, 1
    if Not rsget.Eof then
        prtidx = rsget("prtidx")

        prtidx = format00(4,prtidx)
    end if
    rsget.close
end if

%>
<script language='javascript'>

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

    if (comp.value=="N"){
        resetLimit2Zero();
    }else{
        resetLimit();
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
		  	        //���� ����� 97%�� ���� (��� 10�� �̻��� ��츸) ����
		  	        if (e.dumistock>=10){
		  	            e.value = parseInt(e.dumistock*0.97);
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
	if ((frm.itemrackcode.value.length>0)&&(frm.itemrackcode.value.length!=4)){
		alert('��ǰ ���ڵ�� 4�ڸ��� �����Ǿ��ֽ��ϴ�.');
		frm.itemrackcode.focus();
		return;
	}

    //���� �Ǹ� �Ӽ��� �ٸ� �� ����: ->���� ����O �Ǹ�X ����,
    if (frm.dispyn[1].checked&&frm.sellyn[0].checked){
        alert('���� ���ϴ� ��ǰ�� �Ǹ� �� �� �����ϴ�.');
        frm.dispyn[0].focus();
        return;
    }
<% if (session("ssBctId") <> "icommang") then %>
    if ((frm.dispyn[0].checked&&frm.sellyn[1].checked)||(frm.dispyn[1].checked&&frm.sellyn[0].checked)){
        alert('���� �Ǹ� �Ӽ��� �ٸ��� ���� �� �� �����ϴ�. \n\n���þ���=�Ǹž��� or ������=�Ǹ���');
        frm.dispyn[0].focus();
        return;
    }
<% end if %>

<% if oitem.FResultCount>0 then %>
    <% if Not oitem.FOneItem.IsUpchebeasong then %>
    //(����) �Ǹ� N �ΰ�� ����ǰ�� �Ǵ� MDǰ���� ���� �ؾ���.
    if ((frm.sellyn[1].checked)&&!((frm.danjongyn[2].checked)||(frm.danjongyn[3].checked))){
        alert('�Ǹ� ���� ��ǰ�ΰ�� ����ǰ�� �Ǵ� MDǰ���� �����ϼž� �մϴ�.');
        frm.danjongyn[2].focus();
        return;
    }

    //�Ͻ�ǰ��,���������� �����Ǹ��ΰ�츸 ������
	if ((frm.danjongyn[1].checked)||(frm.danjongyn[2].checked)||(frm.danjongyn[3].checked)){
		if (!frm.limityn[0].checked){
			alert('���� �Ǹ��� ��츸 �Ͻ�ǰ��,����ǰ��, MDǰ���� ���� �� �� �ֽ��ϴ�.');
			frm.limityn[0].focus();
			return;
		}
	}
	<% end if %>
<% end if %>

	//�������̳� �����ϴ°��
	if ((frm.isusing[1].checked)&&(frm.dispyn[0].checked)){
        alert('��� ���� ��ǰ�� �Ǹŷ� ���� �Ұ��մϴ�.');
        frm.dispyn[1].focus();
        return;
    }

	frm.itemoptionarr.value = "";
	//�ɼ� ���� ���� ����
	frm.optremainnoarr.value = "";
	//�ɼ� ��� ����
	frm.optisusingarr.value = "";

    var option_isusing_count = 0;
	for (i = 0; i < frm.elements.length; i++) {
		var e = frm.elements[i];
		if ((e.type == 'text')||(e.type == 'radio')) {
		  	if ((e.name.substring(0,"optremainno".length)) == "optremainno"){
		  	    //���ڸ� ����
		  	    if (!IsDigit(e.value)){
		  	        alert('���� ������ ���ڸ� �����մϴ�.');
		  	        e.select();
		  	        e.focus();
		  	        return;
		  	    }

				frm.itemoptionarr.value = frm.itemoptionarr.value + e.id + "," ;
				frm.optremainnoarr.value = frm.optremainnoarr.value + e.value + "," ;

				if (e.id == "0000") {
				    option_isusing_count = 1;
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
		}
  	}

    if (option_isusing_count < 1) {
        alert("��� �ɼ��� ���������� �Ҽ� �����ϴ�. ��ǰ������ ���������� �����ϰų�, ���þ��� �����ϼ���.");
        return;
    }


	var ret = confirm('���� �Ͻðڽ��ϱ�?');

	if(ret){
		frm.submit();
	}
}

function popoptionEdit(iid){
	var popwin = window.open('/common/pop_adminitemoptionedit.asp?itemid=' + iid,'popitemoptionedit','width=700 height=500 scrollbars=yes resizable=yes');
	popwin.focus();
}

function CloseWindow() {
    window.close();
}

function ReloadWindow() {
    document.location.reload();
}

window.resizeTo(560,700);
</script>

<!-- TOP -->
<table width="280" border="0" align="center" cellpadding="1" cellspacing="1" class="a">
  <tr height="20">
  	<td>
    	<img src="/images/icon_star.gif" align="absbottom"><font color="red"><strong>��ǰ�˻�</strong></font>
	</td>
	<td align="right">
    	<a href="/PDAadmin/index.asp">HOME</a>
	</td>
  </tr>
</table>
<!-- TOP -->

<!-- ǥ ��ܰ˻� ����-->
<table width="280" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor=#BABABA>
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<tr height="25" valign="bottom" bgcolor="F4F4F4">
	        <td valign="top" bgcolor="F4F4F4">
	        	<input type="text" class="text" name="itemid" value="<%= itemid %>" Maxlength="9" size="13">
	        	<input type="button" class="button" value="�˻�">
	        </td>
	</tr>
	</form>
</table>
<!-- ǥ ��ܰ˻� ��-->



<% if oitem.FResultCount>0 then %>
<table width="280" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor=#BABABA>
<form name=frm2 method=post action="do_simpleiteminfoedit.asp">
<input type=hidden name=itemid value="<%= itemid %>">
<input type=hidden name=itemoptionarr value="">
<input type=hidden name=optisusingarr value="">
<input type=hidden name=optremainnoarr value="">

	<tr bgcolor="#FFFFFF">
		<td><%= oitem.FOneItem.Fitemname %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td><%= oitem.FOneItem.Fmakerid %>/<%= oitem.FOneItem.FBrandName %></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td>
			<%= FormatNumber(oitem.FOneItem.Forgprice,0) %>/<%= FormatNumber(oitem.FOneItem.Forgsuplycash,0) %>
			<font color="<%= oitem.FOneItem.getMwDivColor %>"><%= oitem.FOneItem.getMwDivName %></font>
			<% if oitem.FOneItem.FSellcash<>0 then %>
			<%= CLng((1- oitem.FOneItem.Forgsuplycash/oitem.FOneItem.Forgprice)*100) %>%
			<% end if %>
		</td>
	</tr>
	<% if (oitem.FOneItem.FSailYn="Y") then %>
	<tr  bgcolor="#FFFFFF">
		<td>
			<font color="red">
				<%= FormatNumber(oitem.FOneItem.FSellcash,0) %>/<%= FormatNumber(oitem.FOneItem.FBuycash,0) %>
				<% if (oitem.FOneItem.Forgprice<>0) then %>
					����
			        <%= CLng((oitem.FOneItem.Forgprice-oitem.FOneItem.Fsellcash)/oitem.FOneItem.Forgprice*100) %>%
			    <% end if %>

			</font>
			<font color="<%= oitem.FOneItem.getMwDivColor %>"><%= oitem.FOneItem.getMwDivName %></font>
			<% if oitem.FOneItem.FSellcash<>0 then %>
				<%= CLng((1- oitem.FOneItem.FBuycash/oitem.FOneItem.FSellcash)*100) %>%
			<% end if %>
		</td>
	</tr>
	<% end if %>

	<% if (oitem.FOneItem.Fitemcouponyn="Y") then %>
	<tr bgcolor="#FFFFFF">
		<td>
			<font color="green">
				<%= FormatNumber(oitem.FOneItem.GetCouponAssignPrice,0) %>
				&nbsp;&nbsp;
				<%= oitem.FOneItem.GetCouponDiscountStr %> ����
			</font>
		</td>
	</tr>
	<% end if %>

	<tr align="right" bgcolor="#FFFFFF">
		<td>
			��ǰ���ڵ�<input type="text" class="text" name="itemrackcode" value="<%= oitem.FOneItem.FitemRackCode %>" size="4" maxlength="4" >
		</td>
	</tr>
</table>

<table width="280" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor=#BABABA>
	<tr align="center" bgcolor="#FFFFFF">
		<td>����</td>
		<td>�Ǹ�</td>
		<td>���</td>
		<td>����</td>
		<td>����</td>
	</td>
	<tr align="center" bgcolor="#FFFFFF">
		<td><%= oitem.FOneItem.FDispYn %></td>
		<td><%= oitem.FOneItem.FSellYn %></td>
		<td><%= oitem.FOneItem.FIsUsing %></td>
		<td><%= oitem.FOneItem.Fdanjongyn %></td>
		<td><%= oitem.FOneItem.FLimitYn %></td>
	</td>
</table>
<p>

<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor=#BABABA>
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
			<font color="<%= oitem.FOneItem.getMwDivColor %>"><%= oitem.FOneItem.getMwDivName %></font>
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
			<font color="<%= oitem.FOneItem.getMwDivColor %>"><%= oitem.FOneItem.getMwDivName %></font>
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
			<input type="text" name="itemrackcode" value="<%= oitem.FOneItem.FitemRackCode %>" size="4" maxlength="4" > (4�ڸ� Fix)
		</td>
		<td rowspan="4" align="right" bgcolor="#FFFFFF">
			<img src="<%= oitem.FOneItem.FListImage %>" width="100" align="right">
		</td>
	</tr>
	<tr height="25">
		<td bgcolor="#DDDDFF">���ɼ�</td>
		<td bgcolor="#FFFFFF">
		(<%= oitem.FOneItem.FOptionCnt %> ��)
		&nbsp;
		<input type=button class="button" value="�ɼǼ���" onclick="popoptionEdit('<%= itemid %>');">
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
		</table>
		</td>
	</tr>
	<tr>
		<td colspan="2" bgcolor="#FFFFFF">

			<table width="100%" cellspacing=1 cellpadding=1 class=a bgcolor=#BABABA>
			<tr height="25">
				<td width="120" bgcolor="#DDDDFF">��ǰ ���ÿ���</td>
				<td bgcolor="#FFFFFF">
					<% if oitem.FOneItem.FDispYn="Y" then %>
					<input type="radio" name="dispyn" value="Y" checked >������
					<input type="radio" name="dispyn" value="N" >���þ���
					<% else %>
					<input type="radio" name="dispyn" value="Y" >������
					<input type="radio" name="dispyn" value="N" checked ><font color="red">���þ���</font>
					<% end if %>
				</td>
			</tr>
			<tr height="25">
				<td bgcolor="#DDDDFF">��ǰ �Ǹſ���</td>
				<td bgcolor="#FFFFFF">
					<% if oitem.FOneItem.FSellYn="Y" then %>
					<input type="radio" name="sellyn" value="Y" checked >�Ǹ���
					<input type="radio" name="sellyn" value="N" >�Ǹž���
					<% else %>
					<input type="radio" name="sellyn" value="Y" >�Ǹ���
					<input type="radio" name="sellyn" value="N" checked ><font color="red">�Ǹž���</font>
					<% end if %>
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
				<td bgcolor="#DDDDFF">��ǰ ��������</td>
				<td bgcolor="#FFFFFF">
					<% if oitem.FOneItem.Fdanjongyn="Y" then %>
    					<input type="radio" name="danjongyn" value="N" >������
    					<input type="radio" name="danjongyn" value="S" >�Ͻ�ǰ��(7���̻�)
    					<input type="radio" name="danjongyn" value="Y" checked ><font color="red">����ǰ��</font>
    					<input type="radio" name="danjongyn" value="M" >MDǰ��
					<% elseif oitem.FOneItem.Fdanjongyn="S" then %>
    					<input type="radio" name="danjongyn" value="N" >������
    					<input type="radio" name="danjongyn" value="S" checked ><font color="red">�Ͻ�ǰ��(7���̻�)</font>
    					<input type="radio" name="danjongyn" value="Y" >����ǰ��
    					<input type="radio" name="danjongyn" value="M" >MDǰ��
					<% elseif oitem.FOneItem.Fdanjongyn="M" then %>
    					<input type="radio" name="danjongyn" value="N" >������
    					<input type="radio" name="danjongyn" value="S" >�Ͻ�ǰ��(7���̻�)
    					<input type="radio" name="danjongyn" value="Y" >����ǰ��
    					<input type="radio" name="danjongyn" value="M" checked ><font color="red">MDǰ��</font>
					<% else %>
    					<input type="radio" name="danjongyn" value="N" checked >������
    					<input type="radio" name="danjongyn" value="S" >�Ͻ�ǰ��(7���̻�)
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
				<td width="50">�ɼ��ڵ�</td>
				<td>�ɼǸ�</td>
				<td width="100">�ɼǻ�뿩��</td>
				<td width="40">����<br>����</td>
				<td width="80">�����Ǹż���<br><input name="recalcuLimit" type="button" class="button" value="��������" onclick="resetLimit();" <%= chkIIF(oitem.FOneItem.FLimitYn="N","disabled","") %>></td>
				<td width="80"><a href="javascript:TnPopItemStock('<%= itemid %>','');">���������</a></td>
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
							<% if oitemoption.FITemList(i).Foptisusing="Y" then %>
							<input type="radio" name="optisusing<%= oitemoption.FITemList(i).FItemOption %>" value="Y" checked >Y <input type="radio" name="optisusing<%= oitemoption.FITemList(i).FItemOption %>" value="N" >N
							<% else %>
							<input type="radio" name="optisusing<%= oitemoption.FITemList(i).FItemOption %>" value="Y" >Y <input type="radio" name="optisusing<%= oitemoption.FITemList(i).FItemOption %>" value="N" checked ><font color="red">N</font>
							<% end if %>
						</td>
						<td><%= oitemoption.FITemList(i).GetOptLimitEa %></td>
						<td>
							<input type="text" id="<%= oitemoption.FITemList(i).FItemOption %>" dumistock="<%= oitemoption.FITemList(i).GetLimitStockNo %>" name="optremainno<%= oitemoption.FITemList(i).FItemOption %>" value="<%= oitemoption.FITemList(i).GetOptLimitEa %>" size="4" maxlength=5 <% if oitem.FOneItem.FLimitYn="N" then response.write "disabled" %> >
						</td>
						<td <%= chkIIF(oitemoption.FITemList(i).GetLimitStockNo<10,"bgcolor='#6666EE'","") %> ><a href="javascript:TnPopItemStock('<%= itemid %>','<%= oitemoption.FITemList(i).FItemOption %>');"><%= oitemoption.FITemList(i).GetLimitStockNo %></a></td>
					</tr>
				<% next %>
			<% else %>
					<tr align="center" bgcolor="#FFFFFF">
						<td>0000</td>
						<td colspan="2">�ɼǾ���</td>
						<td><%= oitem.FOneItem.GetLimitEa %></td>
						<td>
							<input type="text" id="0000" dumistock="<%= oitem.FOneItem.GetLimitStockNo %>" name="optremainno" value="<%= oitem.FOneItem.GetLimitEa %>" size="4" maxlength=5 <% if oitem.FOneItem.FLimitYn="N" then response.write "disabled" %> >
						</td>
						<td <%= chkIIF(oitem.FOneItem.GetLimitStockNo<10,"bgcolor='#6666EE'","") %> ><a href="javascript:TnPopItemStock('<%= itemid %>','');"><%= oitem.FOneItem.GetLimitStockNo %></a></td>
					</tr>
			<% end if %>
			</table>
		</td>
	</tr>

	<input type=hidden name="pojangok" value="<%= oitem.FOneItem.FPojangOK %>">

</form>
</table>

<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
    <tr valign="top" bgcolor="F4F4F4" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="center" bgcolor="F4F4F4">
          <input type="button" value="�����ϱ�" onclick="SaveItem(frm2)">
          <input type="button" value=" �� �� " onclick="CloseWindow()">
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" bgcolor="F4F4F4" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- ǥ �ϴܹ� ��-->
<% end if %>

<%
set oitemoption = Nothing
set oitem = Nothing

%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->