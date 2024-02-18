<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/baditemcls.asp"-->
<!-- #include virtual="/lib/classes/stock/summary_itemstockcls.asp"-->
<%

dim makerid,mode, searchtype
makerid = request("makerid")
mode = request("mode")
searchtype = request("searchtype")

searchtype = "bad"

dim osummarystock
set osummarystock = new CSummaryItemStock
osummarystock.FRectmakerid = makerid
osummarystock.FRectSearchType = searchtype

if (makerid<>"") then
    osummarystock.GetDailyErrItemListByBrand
else
    osummarystock.GetDailyErrBadItemListByBrandGroup
end if

dim i

%>
<script language='javascript'>
function PopBadItemReInput(makerid){
	var popwin = window.open('/common/pop_baditem_re_input.asp?makerid=' + makerid,'pop_baditem_input','width=900,height=400,resizable=yes,scrollbars=yes')
	popwin.focus();
}

function PopBadItemLossInput(makerid){
	var popwin = window.open('/common/pop_baditem_re_input.asp?makerid=' + makerid + '&actType=actloss','pop_baditem_input','width=900,height=400,resizable=yes,scrollbars=yes')
	popwin.focus();
}

//function getOnLoad(){
//	document.frm.itemid.focus();
//	document.frm.itemid.select();
//}
//
//window.onload=getOnLoad;

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

function SubmitSearchByBrandNew(makerid) {
	document.frm.makerid.value = makerid;
	document.frm.submit();
}

function SubmitSearchByItemId() {
        if (document.frm.itemid.value.length != 12) {
                alert("��ǰ�ڵ带 ��Ȯ�� �Է��ϼ���.");
                return;
        }
        document.frm.makerid.selectedIndex = 0;
        document.frm.submit();
}

function SubmitSearchItem() {
        if (document.frm.itemid.value.length != 12) {
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
        if (document.frm.itemid.value.length != 12) {
                alert("��ǰ�ڵ带 ��Ȯ�� �Է��ϼ���.");
                return;
        }

        var e;
        var t;
        var found = 0;

        e = document.frmlist.elements;
        t = document.frm.itemid.value;
	for (var i=0; i < e.length; i++){
		if (e[i].name == "itemgubun") {
        		if ((e[i].value == t.substring(0,2)) && (e[i+1].value == (1 * t.substring(2,8))) && (e[i+2].value == t.substring(8))) {
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
    }
    <% end if %>
}

function CheckInsert(f, maxvalue){
        alert(f.value);
        if (f.value = "") {
                return;
        }

        if (f.value < 0) {
                alert("�ҷ���ǰ������ ���̳ʽ��� �ɼ� �����ϴ�.");
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
                        alert("�ҷ���ǰ������ ���̳ʽ��� �ɼ� �����ϴ�.");
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

        		        if (pmwdiv==""){
                		    pmwdiv = e[i+5].value;
                		}else{
                		    if (pmwdiv!=e[i+5].value){
                		        alert('���� �Ӽ��� �ٸ���ǰ�� ���� ó�� �� �� �����ϴ�.');
                		        return;
                		    }
                		}
        		}


		}
	}

	if (f.itemgubunarr.value == "") {
        alert("��ǰó���� ��ǰ�� �����ϴ�.");
        return;
    }

    //���ԼӼ��� �ٸ������� ���� �ۼ��� �� ����.


    if (confirm('��ǰ �������� �ۼ��Ͻðڽ��ϱ�?')){
        document.frm.method = "post";
        document.frm.mode.value = "ipgoarr";
        document.frm.pmwdiv.value = pmwdiv;
        document.frm.action = "/common/do_bad_item_re_input.asp";
        document.frm.submit();
    }
}

function ChangePage(v) {
	var frm = document.frm;

	if (v == "bad") {
		frm.action = "baditem_return_list.asp";
	} else {
		frm.action = "erritem_loss_list.asp";
	}

	frm.submit();
}

</script>


<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="1">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			�귣�� : <% drawSelectBoxDesignerwithName "makerid",makerid %>
			&nbsp;
			<input type="radio" name="searchtype" value="bad" <% if (searchtype = "bad") then %>checked<% end if %> onClick="ChangePage('bad')" > �ҷ���ǰ
			<input type="radio" name="searchtype" value="err" <% if (searchtype = "err") then %>checked<% end if %> onClick="ChangePage('err')"> ������ϻ�ǰ
		</td>
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	</form>
</table>
<!-- �˻� �� -->

<p>

<% if makerid<>"" then %>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<input type="button" class="button" value="�ҷ���ǰ��ǰ" onclick="PopBadItemReInput('<%= makerid %>')" border="0">
			&nbsp;
        	<input type="button" class="button" value="�ҷ���ǰ�ν�ó��" onclick="PopBadItemLossInput('<%= makerid %>')" border="0">
		</td>
	</tr>
</table>
<!-- �׼� �� -->

<p>

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			�˻���� : <b><%= osummarystock.FTotalCount %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="30">����</td>
		<td width="50">��ǰ�ڵ�</td>
		<td width="40">�ɼ�</td>
		<td width="50">�̹���</td>
    	<td width="100">�귣��ID</td>

		<td>�����۸�</td>
		<td>�ɼǸ�</td>
		<td width="40">���<br>����</td>

		<td width="50">�Һ��ڰ�</td>
		<td width="40">�ҷ�<br>����</td>
    </tr>

	<% for i=0 to osummarystock.FResultCount - 1 %>
    <tr align="center" bgcolor="#FFFFFF">
    	<td><%= osummarystock.FItemList(i).FItemgubun %></td>
		<td><%= osummarystock.FItemList(i).FItemid %></td>
		<td><%= osummarystock.FItemList(i).FItemoption %></td>
		<td><img src="<%= osummarystock.FItemList(i).Fimgsmall %>" width="50" height="50" onError="this.src='http://webimage.10x10.co.kr/images/no_image.gif'" ></td>
    	<td><%= osummarystock.FItemList(i).Fmakerid %></td>

		<td align="left"><%= osummarystock.FItemList(i).FItemname %></td>
		<td align="left"><%= osummarystock.FItemList(i).FItemOptionName %></td>
		<td><%= osummarystock.FItemList(i).GetMwDivName %></td>

		<td align="right"><%= formatnumber(osummarystock.FItemList(i).Fsellcash,0) %></td>
		<td><%= osummarystock.FItemList(i).Ferrbaditemno %></td>
    </tr>
    <% next %>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
		</td>
	</tr>
</table>

<% else %>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			�˻���� : <b><%= osummarystock.FResultCount %></b>
		</td>
	</tr>
 	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="150">�귣��</td>
		<td width="100">�ҷ���ǰ��On</td>
		<td width="100">�ҷ���ǰ��Off</td>
		<td >&nbsp;</td>
	</tr>
	<% for i=0 to osummarystock.FResultCount-1 %>
	<tr bgcolor="#FFFFFF" height="30">
	    <td><a href="javascript:SubmitSearchByBrandNew('<%= osummarystock.FItemList(i).FMakerid %>');"><%= osummarystock.FItemList(i).FMakerid %></a></td>
	    <td align="center"><%= osummarystock.FItemList(i).FOnCnt %></td>
	    <td align="center"><%= osummarystock.FItemList(i).FOffCnt %></td>
	    <td align="left">
			<input type="button" class="button" value="�ҷ���ǰ��ǰ" onclick="PopBadItemReInput('<%= osummarystock.FItemList(i).FMakerid %>')" border="0">
			&nbsp;
        	<input type="button" class="button" value="�ҷ���ǰ�ν�ó��" onclick="PopBadItemLossInput('<%= osummarystock.FItemList(i).FMakerid %>')" border="0">
	    </td>
	</tr>
	<% next %>
</table>
<% end if %>

<p>




<%
set osummarystock = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
