<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/baditemcls.asp"-->
<%

dim itemid, i, itembarcode

itemid = request("itemid")

dim obaditem

set obaditem = new CBadItem
''obaditem.FRectItemID = itemid

obaditem.GetTempItemList

%>
<script language='javascript'>
function getOnLoad(){
	document.frm.itembarcode.focus();
	document.frm.itembarcode.select();
}

window.onload=getOnLoad;

function SubmitInsert(){
	if (document.frm.itembarcode.value.length <10) {
        alert("��ǰ���ڵ尡 �߸� �ԷµǾ����ϴ�.");
        document.frm.itembarcode.select();
        document.frm.itembarcode.focus();
        return false;
    }

    document.frm.method = "post";
    document.frm.mode.value = "insert";
    document.frm.action = "baditem_input_process.asp";
    document.frm.submit();

    return true;
}

function SubmitDelete(itemgubun, itemid, itemoption){
	if (confirm("�ش��ǰ�� �ҷ���Ͽ��� �����մϴ�. �����Ͻðڽ��ϱ�?") != true) {
        return;
    }
    if (itemid*1>=1000000){
        document.frm.itembarcode.value = "" + itemgubun + ("" + (100000000+1*itemid)).substring(1) + "" + itemoption;
    }else{
        document.frm.itembarcode.value = "" + itemgubun + ("" + (1000000+1*itemid)).substring(1) + "" + itemoption;
    }

	if (document.frm.itembarcode.value.length <10) {
        alert("��ǰ���ڵ尡 �߸� �ԷµǾ����ϴ�.");
        document.frm.itembarcode.select();
    	document.frm.itembarcode.focus();
        return;
    }

    document.frm.method = "post";
    document.frm.mode.value = "delete";
    document.frm.action = "baditem_input_process.asp";
    document.frm.submit();
}

function SubmitModify(f, itemgubun, itemid, itemoption){
	if (confirm("�ش��ǰ�� ������ �����մϴ�. �����Ͻðڽ��ϱ�?") != true) {
        return;
    }
    if (itemid*1>=1000000){
        document.frm.itembarcode.value = "" + itemgubun + ("" + (100000000+1*itemid)).substring(1) + "" + itemoption;
    }else{
        document.frm.itembarcode.value = "" + itemgubun + ("" + (1000000+1*itemid)).substring(1) + "" + itemoption;
    }

	if (document.frm.itembarcode.value.length <10) {
        alert("��ǰ���ڵ尡 �߸� �ԷµǾ����ϴ�.");
        document.frm.itembarcode.select();
        document.frm.itembarcode.focus();
        return;
    }

    document.frm.itemcount.value = f.itemno.value;
    document.frm.method = "post";
    document.frm.mode.value = "modify";
    document.frm.action = "baditem_input_process.asp";
    document.frm.submit();
}

function SubmitList(){
	window.open('/common/pop_item_search.asp','pop_item_search','width=900,height=600,scrollbars=yes');
}


function ReActItems(itemgubunarr,
                    itemarr,
                    itemoptionarr,
                    sellcasharr,
                    suplycasharr,
                    buycasharr,
                    itemnoarr,
                    itemnamearr,
                    itemoptionnamearr,
                    designerarr,
                    mwdivarr)
{
        document.frm.itemgubunarr.value = itemgubunarr;
        document.frm.itemidarr.value = itemarr;
        document.frm.itemoptionarr.value = itemoptionarr;
        document.frm.itemnoarr.value = itemnoarr;

        document.frm.method = "post";
        document.frm.mode.value = "arrinsert";
        document.frm.action = "baditem_input_process.asp";
        document.frm.submit();

        return true;
}





function SubmitUpdateAll(){
        document.frm.method = "post";
        document.frm.mode.value = "tmpbaditem2input";
        document.frm.action = "/admin/stock/dostockrefresh.asp";
        document.frm.submit();
}
</script>



<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method=get action="baditem_input_process.asp" onsubmit="return false;">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="itemcount" value="1">
	<input type="hidden" name="itemgubunarr" value="">
	<input type="hidden" name="itemidarr" value="">
	<input type="hidden" name="itemoptionarr" value="">
	<input type="hidden" name="itemnoarr" value="">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">

			<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
			<tr>
				<td>
					��ǰ���ڵ�:
					<input type=text class="text" name=itembarcode value="<%= itembarcode %>" size=16 maxlength=16 AUTOCOMPLETE="off" onKeyPress="if (event.keyCode == 13){ SubmitInsert(); return false;}">
					<!--
			    	<input type="text" class="text" name="itemid" value="<%= itemid %>" Maxlength="12" size="14" onKeyPress="if (event.keyCode == 13) { SubmitInsert(); return false; }">&nbsp;
			    	-->
			    	<input type="button" class="button" value="�ҷ���ǰ���" onclick="SubmitInsert()">
			    	<input type="button" class="button" value="��ǰ�˻�" onclick="SubmitList()">
				</td>
				<td align="right">
					<input type="button" class="button" value="��ü����" onclick="SubmitUpdateAll()">
				</td>
			</tr>
			</table>

		</td>
	</tr>
	</form>

	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
      <td width="100">�귣��ID</td>
      <td width="60">��۱���</td>
      <td width="60">���Ա���</td>
      <td width="25">����</td>
      <td width="50">��ǰ�ڵ�</td>
      <td>�����۸�</td>
      <td>�ɼ�</td>
      <td width="50">�Һ��ڰ�</td>
      <td width="50">��ϼ���</td>
      <td width="110">-</td>
    </tr>
<% for i=0 to obaditem.FResultCount - 1 %>
    <tr align="center" bgcolor="#FFFFFF">
      <td><%= obaditem.FItemList(i).Fmakerid %></td>
      <td><%= obaditem.FItemList(i).GetdeliverytypeName %></td>
      <td><%= obaditem.FItemList(i).GetMwDivName %></td>
      <td><%= obaditem.FItemList(i).FItemgubun %></td>
      <td><%= obaditem.FItemList(i).FItemid %></td>
      <td align="left"><%= obaditem.FItemList(i).FItemname %></td>
      <td align="left"><%= obaditem.FItemList(i).FItemOptionName %></td>
      <td align="right"><%= formatnumber(obaditem.FItemList(i).Fsellcash,0) %></td>
      <form name=frm<%= i %> onsubmit="return false">
      <td>
        <input type="text" name="itemno" value="<%= obaditem.FItemList(i).Fitemno %>" size="3">
      </td>
      <td>
        <input type="button" class="button" value=" ���� " onclick="SubmitModify(document.frm<%= i %>, '<%= obaditem.FItemList(i).FItemgubun %>', '<%= obaditem.FItemList(i).FItemid %>', '<%= obaditem.FItemList(i).FItemOption %>')">
        <input type="button" class="button" value=" ���� " onclick="SubmitDelete('<%= obaditem.FItemList(i).FItemgubun %>', '<%= obaditem.FItemList(i).FItemid %>', '<%= obaditem.FItemList(i).FItemOption %>')">
      </td>
      </form>
    </tr>
   	<% next %>
<% if obaditem.FResultCount = 0 then %>
    <tr align="center" bgcolor="#FFFFFF">
      <td colspan="10" align="center">�˻��� ��ǰ�� �����ϴ�.</td>
    </tr>
<% end if %>


</table>


<!-- ǥ �ϴܹ� ����-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
    <tr valign="top" bgcolor="F4F4F4" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="right" bgcolor="F4F4F4">&nbsp;</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" bgcolor="F4F4F4" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- ǥ �ϴܹ� ��-->


<%
set obaditem = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->