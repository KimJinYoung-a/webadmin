<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �������� ���������ֹ�������
' History : 2010.06.03 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stock/offshop_balju.asp"-->
<%

menupos = request("menupos")



dim page, shopid, chulgoyn, showdeleted, showmichulgo, michulgoreason
dim statecd, itemid, brandid, shopdiv, baljucode
dim day5chulgo, shortchulgo, tempshort, danjong, etcshort
dim boxno
dim research, i
dim dateType

page = request("page")
shopid = request("shopid")
chulgoyn = request("chulgoyn")
showdeleted = request("showdel")		'������ ������Ʈ�� �Ķ������ delete ������ �ִ� ��� ���´�.
showmichulgo = request("showmichulgo")
michulgoreason = request("michulgoreason")
boxno = request("boxno")

statecd = request("statecd")
itemid = request("itemid")
brandid = request("brandid")
shopdiv = request("shopdiv")
baljucode = request("baljucode")

day5chulgo = request("day5chulgo")
shortchulgo = request("shortchulgo")
tempshort = request("tempshort")
danjong = request("danjong")
etcshort = request("etcshort")


research = request("research")
dateType = requestCheckVar(request("dateType"),1)

if (page = "") then
	page = 1
end if


if (research = "") then
	showdeleted = "N"
	michulgoreason = "all"
end if



michulgoreason = "|"
if (day5chulgo = "Y") then
	'5�ϳ����
	michulgoreason = michulgoreason + "5|"
end if
if (shortchulgo = "Y") then
	'������
	michulgoreason = michulgoreason + "S|"
end if
if (tempshort = "Y") then
	'�Ͻ�ǰ��
	michulgoreason = michulgoreason + "T|"
end if
if (danjong = "Y") then
	'����
	michulgoreason = michulgoreason + "D|"
end if
if (etcshort = "Y") then
	'��Ÿ
	michulgoreason = michulgoreason + "E|"
end if



dim yyyy1,mm1 , dd1, yyyy2, mm2, dd2, fromDate, toDate
yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")
yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")

if (yyyy1="") then
	yyyy1 = Cstr(Year(now()))
	mm1 = Cstr(Month(now()))
	dd1 = Cstr(day(now()))
end if

if (yyyy2="") then
	yyyy2 = Cstr(Year(now()))
	mm2 = Cstr(Month(now()))
	dd2 = Cstr(day(now()))
end if

fromDate = DateSerial(yyyy1, mm1, dd1)
toDate = DateSerial(yyyy2, mm2, dd2+1)



dim oshopbalju

set oshopbalju = new CShopBalju

oshopbalju.FRectFromDate = fromDate
oshopbalju.FRectToDate = toDate
oshopbalju.FRectDateType = dateType
oshopbalju.FRectBaljuId = shopid

oshopbalju.FRectItemid = itemid
oshopbalju.FRectBrandid = brandid
oshopbalju.FRectShopdiv = shopdiv
oshopbalju.FRectBaljucode = baljucode
oshopbalju.FRectBoxno = boxno

if (statecd = "A") then
	oshopbalju.FRectChulgoYN = "N"
else
	oshopbalju.FRectStatecd = statecd
end if

oshopbalju.FRectShowDeleted = "N"
oshopbalju.FRectMichulgoReason = michulgoreason

oshopbalju.FCurrPage = page
oshopbalju.Fpagesize = 25

oshopbalju.GetShopBaljuByItem

%>

<script language='javascript'>

function MakeJumun(){
	location.href="jumuninput.asp";
}

function PopSegumil(frm,iidx,comp){
	if (calendarOpen2(comp)){
		if (confirm('������ : ' + comp.value + ' OK?')){
			frm.idx.value = iidx;
			frm.mode.value = "segumil";
			frm.submit();
		}
	};
}

function PopIpgumil(frm,iidx,comp){
	if (calendarOpen2(comp)){
		if (confirm('�Ա��� : ' + comp.value + ' OK?')){
			frm.idx.value = iidx;
			frm.mode.value="ipkumil";
			frm.submit();
		}
	};
}

function PopIpgoSheet(v){
	var popwin;
	popwin = window.open('popshopjumunsheet2.asp?idx=' + v ,'shopjumunsheet','width=740,height=600,scrollbars=yes,status=no');
	popwin.focus();
}

function ExcelSheet(v){
	window.open('popshopjumunsheet2.asp?idx=' + v + '&xl=on');
}

function MakeReJumun(iidx){
	if (!calendarOpen2(frmMaster.datestr)){ return };

	if (!confirm('��� ������ : ' + frmMaster.datestr.value + ' OK?')){ return };

	if (confirm('�̹�� �ֹ����� �ۼ� �Ͻðڽ��ϱ�?')){
		frmMaster.idx.value = iidx;
		frmMaster.mode.value = "remijumun";
		frmMaster.target = "_blank";
		frmMaster.submit();
	}
}

function MakeReturn(iidx){
	if (!calendarOpen2(frmMaster.datestr)){ return };

	if (!confirm('��� ������ : ' + frmMaster.datestr.value + ' OK?')){ return };

	if (confirm('��ǰ �ֹ����� �ۼ� �Ͻðڽ��ϱ�?')){
		frmMaster.idx.value = iidx;
		frmMaster.mode.value = "returnjumun";
		frmMaster.target = "_blank";
		frmMaster.submit();
	}
}

function MakeDuplicateJumun(iidx){
	if (!calendarOpen2(frmMaster.datestr)){ return };

	if (!confirm('��� ������ : ' + frmMaster.datestr.value + ' OK?')){ return };

	if (confirm('���� �ֹ����� �ۼ� �Ͻðڽ��ϱ�?')){
		frmMaster.idx.value = iidx;
		frmMaster.mode.value = "duplicatejumun";
		frmMaster.target = "_blank";
		frmMaster.submit();
	}
}

function Popbalju(){
	var frm = document.frmlist;
	var idxarr = "";
	for (var i=0;i<frm.elements.length;i++){
		if ((frm.elements[i].name=="ck_all") && (frm.elements[i].checked)){
        	idxarr = idxarr + frm.elements[i].value + ",";
      	}
	}
	if (idxarr==""){
		alert('�ֹ����� �����ϼ���.');
		return;
	}else{
		frm.idxarr.value= idxarr;
		frm.target="_blank";
		frm.action="popoffbaljulist.asp"
		frm.submit();
	}
}
</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			ShopID : 
			<% 'drawSelectBoxOffShop "shopid",shopid %>
			<% Call NewDrawSelectBoxDesignerwithNameAndUserDIV("shopid",shopid, "21") %>
			&nbsp;
			<select class="select" name="dateType">
				<option value="B" <%= CHKIIF(dateType="B", "selected", "") %> >������</option>
				<option value="C" <%= CHKIIF(dateType="C", "selected", "") %> >�����</option>
				<option value="J" <%= CHKIIF(dateType="J", "selected", "") %> >�ֹ���</option>
			</select> :
			<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
			&nbsp;
			�ֹ����� :
			<select name="statecd" class="select">
				<option value="">��ü
				<option value=" " <% if statecd=" " then response.write "selected" %> >�ۼ���
				<option value="0" <% if statecd="0" then response.write "selected" %> >�ֹ�����
				<option value="1" <% if statecd="1" then response.write "selected" %> >�ֹ�Ȯ��
				<option value="2" <% if statecd="2" then response.write "selected" %> >�Աݴ��
				<option value="5" <% if statecd="5" then response.write "selected" %> >����غ�
				<option value="6" <% if statecd="6" then response.write "selected" %> >�����
				<option value="7" <% if statecd="7" then response.write "selected" %> >���Ϸ�
				<option value="8" <% if statecd="8" then response.write "selected" %> >�԰���
				<option value="9" <% if statecd="9" then response.write "selected" %> >�԰�Ϸ�
				<option value="">========
				<option value="A" <% if statecd="A" then response.write "selected" %> >���������ü
			</select>
		</td>
		<td rowspan="3" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
			�ֹ��ڵ� : <input type="text" class="text" name="baljucode" value="<%= baljucode %>" size="10" maxlength="8">
			&nbsp;
			�귣�� : <% drawSelectBoxDesignerwithName "brandid", brandid %>
			&nbsp;
			��ǰ�ڵ� : <input type="text" class="text" name="itemid" value="<%= itemid %>" size="10" maxlength="12">
			&nbsp;
			�ڽ���ȣ : <input type="text" class="text" name="boxno" value="<%= boxno %>" size="4" maxlength="12">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
	     	���� SHOP���� :
	     	<input type="radio" name="shopdiv" value="" <% if shopdiv="" then response.write "checked" %> >��ü
			<input type="radio" name="shopdiv" value="direct" <% if shopdiv="direct" then response.write "checked" %> >����
			<input type="radio" name="shopdiv" value="franchisee" <% if shopdiv="franchisee" then response.write "checked" %> >������
			<input type="radio" name="shopdiv" value="foreign" <% if shopdiv="foreign" then response.write "checked" %> >�ؿ�
			<input type="radio" name="shopdiv" value="buy" <% if shopdiv="buy" then response.write "checked" %> >����
			&nbsp;&nbsp;
			|
			&nbsp;&nbsp;
			�������� :
			<input type="checkbox" name="day5chulgo" value="Y" <% if day5chulgo="Y" then response.write "checked" %> >5�ϳ����
			<input type="checkbox" name="shortchulgo" value="Y" <% if shortchulgo="Y" then response.write "checked" %> >������
			<input type="checkbox" name="tempshort" value="Y" <% if tempshort="Y" then response.write "checked" %> >�Ͻ�ǰ��
			<input type="checkbox" name="danjong" value="Y" <% if danjong="Y" then response.write "checked" %> >����
			<input type="checkbox" name="etcshort" value="Y" <% if etcshort="Y" then response.write "checked" %> >��Ÿ
		</td>
	</tr>
	</form>
</table>
<!-- �˻� �� -->

<p>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="16">
			�˻���� : <b><%= oshopbalju.FTotalCount %></b>
			&nbsp;
			������ : <b><%= page %> / <%= oshopbalju.FTotalpage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td>�ֹ���</td>
		<td>������</td>
		<td>�����̵�</td>
		<td>�ڽ�<br>��ȣ</td>
		<td>�����ڵ�</td>
		<td>�ֹ��ڵ�</td>
		<td>������</td>
		<td>�귣��</td>
		<td>�̹���</td>
		<td>����</td>
		<td>��ǰ�ڵ�</td>
		<td>�ɼ�</td>
		<td>��ǰ��(��)<br><font color="blue">[�ɼǸ�(��)]</font></td>
		<td>�ֹ�<br>����</td>
		<td>����<br>����</td>
		<td>���<br>����</td>
		<td>���</td>
	</tr>
	<% if oshopbalju.FResultCount >0 then %>
	<% for i=0 to oshopbalju.FResultcount-1 %>
	<tr bgcolor="#FFFFFF">
		<td align="center"><%= oshopbalju.FItemList(i).Fregdate %></td>
		<td align="center"><%= oshopbalju.FItemList(i).Fbaljudate %></td>
		<td align="center"><%= oshopbalju.FItemList(i).Fbaljuid %><br><%= oshopbalju.FItemList(i).Fbaljuname %></td>
		<td align="center">
			<%
			if (oshopbalju.FItemList(i).Fboxno <> "0") then
				response.write oshopbalju.FItemList(i).Fboxno
			end if
			%>
		</td>
		<td align="center"><%= oshopbalju.FItemList(i).Fbaljunum %></td>
		<td align="center"><%= oshopbalju.FItemList(i).Fbaljucode %></td>
		<td align="center">
			<font color="<%= oshopbalju.FItemList(i).GetStateColor %>"><%= oshopbalju.FItemList(i).GetStateName %></font>
			<% if (oshopbalju.FItemList(i).Frealitemno > 0) then %>
				<br><%= oshopbalju.FItemList(i).FAlinkCode %>
			<% end if %>
		</td>
		<td align="center"><%= oshopbalju.FItemList(i).Fmakerid %></td>
		<td align="center"><img src="<%= oshopbalju.FItemList(i).Fmainimageurl %>" width="50"></td>
		<td align="center"><%= oshopbalju.FItemList(i).Fitemgubun %></td>
		<td align="center"><%= oshopbalju.FItemList(i).Fitemid %></td>
		<td align="center"><%= oshopbalju.FItemList(i).Fitemoption %></td>
		<td align="left">
			<%= oshopbalju.FItemList(i).Fitemname %>
			<% if (oshopbalju.FItemList(i).Fitemoption <> "0000") then %>
				<br><font color="blue">[<%= oshopbalju.FItemList(i).Fitemoptionname %>]</font>
			<% end if %>
		</td>
		<td align="center"><%= oshopbalju.FItemList(i).Fbaljuitemno %></td>
		<td align="center"><%= oshopbalju.FItemList(i).Frealbaljuitemno %></td>
		<td align="center"><%= oshopbalju.FItemList(i).Frealitemno %></td>
		<td align="center">
			<%= oshopbalju.FItemList(i).Fcomment %>
			<%= oshopbalju.FItemList(i).Fipgoflag %>
		</td>
	</tr>
	<% next %>
	<% else %>
<tr bgcolor="#FFFFFF">
		<td colspan="17" align=center>[ �˻������ �����ϴ�. ]</td>
	</tr>
	<% end if %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="17" align="center">
			<%
			dim strparam
			strparam = "&shopid=" + CStr(shopid) + "&yyyy1=" + CStr(yyyy1) + "&mm1=" + CStr(mm1) + "&dd1=" + CStr(dd1) + "&yyyy2=" + CStr(yyyy2) + "&mm2=" + CStr(mm2) + "&dd2=" + CStr(dd2)

			strparam = strparam + "&menupos=" + CStr(menupos)
			strparam = strparam + "&chulgoyn=" + CStr(chulgoyn)
			strparam = strparam + "&showdel=" + CStr(showdeleted)
			strparam = strparam + "&showmichulgo=" + CStr(showmichulgo)
			strparam = strparam + "&michulgoreason=" + Server.URLEncode(CStr(michulgoreason))

			strparam = strparam + "&statecd=" + CStr(statecd)
			strparam = strparam + "&itemid=" + CStr(itemid)
			strparam = strparam + "&brandid=" + CStr(brandid)
			strparam = strparam + "&shopdiv=" + CStr(shopdiv)
			strparam = strparam + "&baljucode=" + CStr(baljucode)

			strparam = strparam + "&day5chulgo=" + CStr(day5chulgo)
			strparam = strparam + "&shortchulgo=" + CStr(shortchulgo)
			strparam = strparam + "&tempshort=" + CStr(tempshort)
			strparam = strparam + "&danjong=" + CStr(danjong)
			strparam = strparam + "&etcshort=" + CStr(etcshort)

			strparam = strparam + "&boxno=" + CStr(boxno)
			%>
			<% if oshopbalju.HasPreScroll then %>
				<a href="?page=<%= oshopbalju.StartScrollPage-1 %>&research=on<%= strparam %>">[pre]</a>
			<% else %>
				[pre]
			<% end if %>

			<% for i=0 + oshopbalju.StartScrollPage to oshopbalju.FScrollCount + oshopbalju.StartScrollPage - 1 %>
				<% if i>oshopbalju.FTotalpage then Exit for %>
				<% if CStr(page)=CStr(i) then %>
				<font color="red">[<%= i %>]</font>
				<% else %>
				<a href="?page=<%= i %>&research=on<%= strparam %>">[<%= i %>]</a>
				<% end if %>
			<% next %>

			<% if oshopbalju.HasNextScroll then %>
				<a href="?page=<%= i %>&research=on<%= strparam %>">[next]</a>
			<% else %>
				[next]
			<% end if %>
		</td>
	</tr>
</table>


<%
set oshopbalju = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
