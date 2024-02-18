<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/baditemcls.asp"-->
<!-- #include virtual="/lib/classes/stock/summary_itemstockcls.asp"-->
<%

dim itemid, makerid, mode, actType

itemid  = requestCheckVar(request("itemid"),32)     '' length > 9
makerid = requestCheckVar(request("makerid"),32)
mode    = requestCheckVar(request("mode"),9)
actType = requestCheckVar(request("actType"),9)     '' actType="actloss" �ν�ó�� actType<>"actloss" ��ǰó��

dim osummarystock
set osummarystock = new CSummaryItemStock
if (Len(itemid) = 12) then
        osummarystock.FRectItemID =  Mid(itemid,3,6)
end if
if (Len(itemid) = 14) then
        osummarystock.FRectItemID =  Mid(itemid,3,8)
end if
osummarystock.FRectmakerid = makerid

if (makerid<>"") then
osummarystock.GetDailyErrItemListByBrand
end if

if (osummarystock.FResultCount > 0) then
        makerid = osummarystock.FItemList(0).Fmakerid
end if

dim i

dim LorRText
if (actType="actloss") then
    LorRText = "�ν�"
else
    LorRText = "��ǰ"
end if

%>
<script language='javascript'>


function getOnLoad(){
	document.frm.itemid.focus();
	document.frm.itemid.select();
}

window.onload=getOnLoad;

function checkhL(e){
    if (e.value*1>0){
        hL(e);
    }else{
        dL(e);
    }
}

function SubmitSearchByBrand() {
        if (document.frm.makerid.selectedIndex == 0) {
                alert("�귣�带 �����ϼ���.");
                return;
        }
        document.frm.itemid.value = "";
        document.frm.submit();
}

function SubmitSearchByItemId() {
        if ((document.frm.itemid.value.length != 12) && (document.frm.itemid.value.length != 14)) {
                alert("��ǰ�ڵ带 ��Ȯ�� �Է��ϼ���.");
                return;
        }
        document.frm.makerid.selectedIndex = 0;
        document.frm.submit();
}

function SubmitSearchItem() {
        if ((document.frm.itemid.value.length != 12) && (document.frm.itemid.value.length != 14)) {
                if (document.frm.makerid.selectedIndex == 0) {
                        alert("�귣�� �Ǵ� ��ǰ�ڵ带 �Է��ϼ���.");
                        return;
                }
                document.frm.itemid.value = "";
                document.frm.submit();
        } else {
                document.frm.makerid.selectedIndex = 0;
                document.frm.submit();
        }
}

function SubmitInsert(){
    <% if (osummarystock.FResultCount < 1) then %>
        SubmitSearchByItemId();
    <% else %>
        if ((document.frm.itemid.value.length != 12) && (document.frm.itemid.value.length != 14)) {
                alert("��ǰ�ڵ带 ��Ȯ�� �Է��ϼ���.");
                return;
        }

        var e;
        var t;
    	var found = 0;
		var itemgubun = "";
		var itemid = "";
		var itemoption = "";

        e = document.frmlist.elements;
        t = document.frm.itemid.value;
	for (var i=0; i < e.length; i++){
		if (e[i].name == "itemgubun") {
			if (e[i].value != "10"){
				alert("���� ��ǰ�� ó�� �� �� �����ϴ�.");
				return;
			}

			if (t.length == 12) {
				itemgubun = t.substring(0,2);
				itemid = (1 * t.substring(2,8));
				itemoption = t.substring(8);
			} else if (t.length == 14) {
				itemgubun = t.substring(0,2);
				itemid = (1 * t.substring(2,10));
				itemoption = t.substring(10);
			} else {
				alert("ERROR");
				return;
			}

			if ((e[i].value == itemgubun) && (e[i+1].value == itemid) && (e[i+2].value == itemoption)) {
				e[i+3].value = (1 * e[i+3].value) + 1;

				if ((1 * e[i+3].value) > (e[i+4].value * -1)) {
						e[i+3].value = (e[i+4].value * -1);
						alert("�ҷ���ϵ� �������� ������ Ů�ϴ�. ���� �ҷ������ �ϼ���.");
							}

				found = 1;
				hL(e[i+3]);
				break;
			}
		}
	}

	if (found == 0) {
        alert("��ǰ�� ��Ͽ� �����ϴ�. �ٸ� �귣���̰ų�, �ҷ������ �Ǿ� ���� �ʽ��ϴ�.");
    }else{
        frm.itemid.select();
        frm.itemid.focus();
    }
    <% end if %>
}

function CheckInsert(f, maxvalue){
        alert(f.value);
        if (f.value = "") {
                return;
        }

        if (f.value < 0) {
                alert("�ҷ�<%= LorRText %>������ ���̳ʽ��� �ɼ� �����ϴ�.");
                f.value = 0;
                return;
        }

        if (f.value > (maxvalue * -1)) {
                alert("�ҷ���ϵ� �������� ������ Ů�ϴ�. ���� �ҷ������ �ϼ���.");
                f.value = (maxvalue * -1);
                return;
        }


}

function SubmitList(){
	window.open('/common/pop_item_search.asp','pop_item_search','width=900,height=600');
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
        document.frm.action = "do_bad_item_input.asp";
        document.frm.submit();

        return true;
}





function SubmitUpdateAll(){
    var e;
    var f;
    var found = 0;
    var pmwdiv = "";

    e = document.frmlist.elements;
    f = document.frm;

    f.itemgubunarr.value = "";
    f.itemidarr.value = "";
    f.itemoptionarr.value = "";
    f.itemnoarr.value = "";

	for (var i=0; i < e.length; i++){
		if (e[i].name == "itemgubun") {
        		if ((e[i+3].value * 0) != 0) {
    		        alert("������ �߸� �ԷµǾ����ϴ�.");
    		        e[i+3].focus();
    		        e[i+3].select();
    		        return;
                }

                if (e[i+3].value == "") {
                        alert("������ �߸� �ԷµǾ����ϴ�.");
    		        e[i+3].focus();
    		        e[i+3].select();
                        return;
                }

                if (e[i+3].value < 0) {
                        alert("�ҷ�<%= LorRText %>������ ���̳ʽ��� �ɼ� �����ϴ�.");
    		        e[i+3].focus();
    		        e[i+3].select();
                        return;
                }

                if (e[i+3].value > (e[i+4].value * -1)) {
                        alert("�ҷ���ϵ� �������� ������ Ů�ϴ�. ���� �ҷ������ �ϼ���.");
    		        e[i+3].focus();
    		        e[i+3].select();
                        return;
                }


        		if ((e[i+3].value * 1) != 0) {
        		        f.itemgubunarr.value = f.itemgubunarr.value + e[i].value + "|";
        		        f.itemidarr.value = f.itemidarr.value + e[i+1].value + "|";
        		        f.itemoptionarr.value = f.itemoptionarr.value + e[i+2].value + "|";
        		        f.itemnoarr.value = f.itemnoarr.value + e[i+3].value + "|";

        		        <% if (actType<>"actloss") then %>
        		        if (pmwdiv==""){
                		    pmwdiv = e[i+5].value;
                		}else{
                		    if (pmwdiv!=e[i+5].value){
                		        alert('���� �Ӽ��� �ٸ���ǰ�� ���� ó�� �� �� �����ϴ�.');
                		        return;
                		    }
                		}
                		<% end if %>
        		}


		}
	}

	if (f.itemgubunarr.value == "") {
        alert("<%= LorRText %>ó���� ��ǰ�� �����ϴ�.");
        return;
    }

    //���ԼӼ��� �ٸ������� ���� �ۼ��� �� ����.


    if (confirm('<%= LorRText %> �������� �ۼ��Ͻðڽ��ϱ�?')){
        document.frm.method = "post";
        <% if (actType="actloss") then %>
        document.frm.mode.value = "lossarr";
        <% else %>
        document.frm.mode.value = "ipgoarr";
        <% end if %>
        document.frm.pmwdiv.value = pmwdiv;
        document.frm.action = "/common/baditem_re_input_process.asp";
        document.frm.submit();
    }
}
</script>

<!-- �˻� ���� -->

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" >
<tr bgcolor="#FFFFFF">
    <td>** <%= chkIIF(actType="actloss"," �ҷ� ��ǰ <strong>�ν�</strong> ó�� "," �ҷ� ��ǰ <strong>��ǰ</strong> ó�� ") %></td>
</tr>
</table>
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method=get action="" onsubmit="return false;">
	<input type="hidden" name="actType" value="<%= actType %>">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="brandid" value="<%= makerid %>">
	<input type="hidden" name="itemgubunarr" value="">
	<input type="hidden" name="itemidarr" value="">
	<input type="hidden" name="itemoptionarr" value="">
	<input type="hidden" name="itemnoarr" value="">
	<input type="hidden" name="pmwdiv" value="">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			�귣��� : <% drawSelectBoxDesignerwithName "makerid",makerid  %>
			<input type="button" class="button_s" value="�귣��ҷ���ǰ��ϰ˻�" onClick="SubmitSearchItem()">
		</td>
	</tr>
</table>

<p>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td>
			��ǰ�ڵ� : <input type="text" class="text" name="itemid" value="<%= itemid %>" Maxlength="14" size="14" onKeyPress="if (event.keyCode == 13) { SubmitInsert(); return false; }">
			<input type="button" class="button" value="<%= chkIIF(actType="actloss"," �ν��߰� "," ��ǰ�߰� ") %>" onclick="SubmitInsert()">
			<!--
			&nbsp;
        	<input type="button" value=" �귣��˻� " onclick="SubmitSearchByBrand()">&nbsp;&nbsp;<input type="button" value=" ��ǰ�ڵ�˻� " onclick="SubmitSearchByItemId()"><br>
        	-->
        	&nbsp;
			* �������� �귣�常 �ϰ���ǰ�԰� �����մϴ�.
		</td>
		<td align="right">
			<input type="button" class="button" value="��ü����" onclick="SubmitUpdateAll()">
		</td>
	</tr>
	</form>
</table>
<!-- �׼� �� -->

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
      <td width="100">�귣��ID</td>
      <td width="40">����<br>����</td>
      <td width="25">����</td>
      <td width="40">��ǰ<br>�ڵ�</td>
      <td width="30">�ɼ�</td>
      <td>�����۸�</td>
      <td>�ɼǸ�</td>
      <td width="50">�Һ��ڰ�</td>
      <td width="40">�ҷ�<br>����</td>
      <td width="40">��ǰ<br>����</td>
    </tr>
    <form name="frmlist" method=get action="" onsubmit="return false;">
<% for i=0 to osummarystock.FResultCount - 1 %>
    <tr align="center" bgcolor="<%= chkIIF(osummarystock.FItemList(i).FItemgubun="10","#FFFFFF","#EEEEEE") %>">
      <td><%= osummarystock.FItemList(i).Fmakerid %></td>
      <td>
        <% if osummarystock.FItemList(i).FItemgubun="10" then %>
        <font color="<%= mwdivColor(osummarystock.FItemList(i).FMwdiv) %>"><%= osummarystock.FItemList(i).GetMwDivName %></font>
        <% end if %>
      </td>
      <td><%= osummarystock.FItemList(i).FItemgubun %></td>
      <td><%= osummarystock.FItemList(i).FItemid %></td>
      <td><%= osummarystock.FItemList(i).FItemoption %></td>
      <td align="left"><%= osummarystock.FItemList(i).FItemname %></td>
      <td align="left"><%= osummarystock.FItemList(i).FItemOptionName %></td>
      <td align="right"><%= formatnumber(osummarystock.FItemList(i).Fsellcash,0) %></td>
      <td>
        <%= osummarystock.FItemList(i).Ferrbaditemno %>
      </td>
      <input type="hidden" name="itemgubun" value="<%= osummarystock.FItemList(i).FItemgubun %>">
      <input type="hidden" name="itemid" value="<%= osummarystock.FItemList(i).FItemid %>">
      <input type="hidden" name="itemoption" value="<%= osummarystock.FItemList(i).FItemOption %>">
      <td>
        <input type="text" class="text" name="itemno" value="0" size="3" onKeyUP="checkhL(this);" <%= chkIIF(osummarystock.FItemList(i).FItemgubun="10","","disabled") %> >
      </td>
      <input type="hidden" name="itemmaxno" value="<%= osummarystock.FItemList(i).Ferrbaditemno %>" >
      <input type="hidden" name="mwdiv" value="<%= osummarystock.FItemList(i).FMwdiv %>">
    </tr>
   	<% next %>
<% if osummarystock.FResultCount = 0 then %>
    <tr align="center" bgcolor="#FFFFFF">
      <td colspan="11" align="center">�˻��� ��ǰ�� �����ϴ�.</td>
    </tr>
<% end if %>
    </form>
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
set osummarystock = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
