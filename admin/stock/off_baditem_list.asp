<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �������� ���������Է·α�
' History : 2009.04.07 �̻� ����
'			2010.04.02 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshop_summary.asp"-->

<%
dim shopid, makerid, errType, itembarcode, itemgubun, itemid, itemoption

shopid       = RequestCheckVar(request("shopid"),32)
makerid      = RequestCheckVar(request("makerid"),32)
itembarcode  = RequestCheckVar(request("itembarcode"),32)
errType      = requestCheckVar(request("errType"),9)

if (itembarcode <> "") then
    if Not (fnGetItemCodeByPublicBarcode(itembarcode,itemgubun,itemid,itemoption)) then
        if Len(itembarcode)=12 then
            itemgubun   = Left(itembarcode, 2)
            itemid      = CStr(Mid(itembarcode, 3, 6) + 0)
            itemoption  = Right(itembarcode, 4)
        elseif Len(itembarcode)=14 then
            itemgubun   = Left(itembarcode, 2)
            itemid      = CStr(Mid(itembarcode, 3, 8) + 0)
            itemoption  = Right(itembarcode, 4)
        end if
    end if
end if

''rw itemgubun & "," & itemid & "," & itemoption
dim yyyy1,mm1,dd1,yyyy2,mm2,dd2
dim fromdate,todate

fromdate    = requestCheckVar(request("fromdate"),10)
todate      = requestCheckVar(request("todate"),10)

if fromdate<>"" then
	yyyy1 = Left(fromdate,4)
	mm1 = Mid(fromdate,6,2)
	dd1 = Mid(fromdate,9,2)
else
	yyyy1 = request("yyyy1")
	mm1 = request("mm1")
	dd1 = request("dd1")
end if

if todate<>"" then
	yyyy2 = Left(todate,4)
	mm2 = Mid(todate,6,2)
	dd2 = Mid(todate,9,2)
else
	yyyy2 = request("yyyy2")
	mm2 = request("mm2")
	dd2 = request("dd2")
end if



if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = Cstr(day(now()))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

fromdate = CStr(DateSerial(yyyy1, mm1, dd1))
todate = CStr(DateSerial(yyyy2, mm2, dd2+1))


dim oOffStock
set oOffStock = new CShopItemSummary
oOffStock.FRectShopID   = shopid
oOffStock.FRectMakerID  = makerid
oOffStock.FRectItemGubun   = itemgubun
oOffStock.FRectItemID   = itemid
oOffStock.FRectItemOption   = itemoption

oOffStock.FRectErrType  = errType
oOffStock.FRectStartDate = fromdate
oOffStock.FRectEndDate   = todate

oOffStock.GetOFFDailyErrItemList

dim i
dim TotErrrealcheckno, TotErrbaditemno, TotErrsampleitemno
TotErrrealcheckno  = 0
TotErrbaditemno    = 0
TotErrsampleitemno = 0

%>
<script language='javascript'>
function EditErrDetail(yyyymmdd,itemgubun,itemid,itemoption,shopid,i){
    var frm = document.frmList;
    var errrealcheckno, errsampleitemno;
    if (frm.errrealcheckno){
        if (frm.errrealcheckno.length){
            errrealcheckno = frm.errrealcheckno[i].value;
        }else{
            errrealcheckno = frm.errrealcheckno.value;
        }
    }
    if (!IsInteger(errrealcheckno)){
        alert('���ڸ� �����մϴ�.');
        return;
    }

    if (frm.errsampleitemno){
        if (frm.errsampleitemno.length){
            errsampleitemno = frm.errsampleitemno[i].value;
        }else{
            errsampleitemno = frm.errsampleitemno.value;
        }
    }
    if (!IsInteger(errsampleitemno)){
        alert('���ڸ� �����մϴ�.');
        return;
    }

    if (confirm('���� �Ͻðڽ��ϱ�?')){
        document.frmDel.mode.value= "OffErrEdit";
        document.frmDel.yyyymmdd.value = yyyymmdd;
        document.frmDel.itemgubun.value = itemgubun;
        document.frmDel.itemid.value = itemid;
        document.frmDel.itemoption.value = itemoption;
        document.frmDel.shopid.value = shopid;
        document.frmDel.errrealcheckno.value = errrealcheckno;
		document.frmDel.errsampleitemno.value = errsampleitemno;
        document.frmDel.submit();
    }
}

function DelErrDetail(yyyymmdd,itemgubun,itemid,itemoption,shopid){
    if (confirm('���� �Ͻðڽ��ϱ�?')){
        document.frmDel.mode.value= "OffErrDelete";
        document.frmDel.yyyymmdd.value = yyyymmdd;
        document.frmDel.itemgubun.value = itemgubun;
        document.frmDel.itemid.value = itemid;
        document.frmDel.itemoption.value = itemoption;
        document.frmDel.shopid.value = shopid;
        document.frmDel.submit();
    }
}

function inputDummiErr(comp,shopid,itemgubun,itemid,itemoption){
    var bufFrm = comp.form;
    var frm = document.frmDel;
    if (bufFrm.bufYYYYMMDD.value.length!=10){
        alert('���� ��¥�� �Է��ϼ���.');
        return;
    }

    if (confirm('������ �޴� - �Է� �Ͻðڽ��ϱ�?')){
        frm.mode.value="dummidailyerrlogOFF";
        frm.yyyymmdd.value=bufFrm.bufYYYYMMDD.value;
        frm.shopid.value=shopid;
		frm.itemgubun.value=itemgubun;
        frm.itemid.value=itemid;
        frm.itemoption.value=itemoption;
        frm.errrealcheckno.value=0;

        frm.submit();
    }
}

function inputDummiWithLstErrNo(comp,shopid,itemgubun,itemid,itemoption,preyyyymm){
    var bufFrm = comp.form;
    var frm = document.frmDel;
    if (bufFrm.bufYYYYMMDD.value.length!=10){
        alert('���� ��¥�� �Է��ϼ���.');
        return;
    }

    if (confirm('������ �޴� - �Է� �Ͻðڽ��ϱ�?')){
        frm.mode.value="dummidailyerrlogCHGOFF";
        frm.yyyymmdd.value=bufFrm.bufYYYYMMDD.value;
        frm.preyyyymmdd.value=preyyyymm;
        frm.shopid.value=shopid;
		frm.itemgubun.value=itemgubun;
        frm.itemid.value=itemid;
        frm.itemoption.value=itemoption;
        frm.errrealcheckno.value=bufFrm.bufNo.value;

        frm.submit();
    }
}

function popShopCurrentStock(shopid,itemgubun,itemid,itemoption){
    var popwin = window.open('/common/offshop/shop_itemcurrentstock.asp?shopid=' + shopid + '&itemgubun=' + itemgubun + '&itemid=' + itemid + '&itemoption=' + itemoption,'popShopCurrentStock','width=900,height=600,resizable=yes,scrollbars=yes');
    popwin.focus();
}

function submitFrm() {
	document.frm.submit();
}

</script>
<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name=frm method=get>
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
		    ���� : <% drawSelectBoxOffShop "shopid",shopid %>
		    &nbsp;
        	�귣��� : <% drawSelectBoxDesignerwithName "makerid",makerid  %>
        	&nbsp;
			��ǰ���ڵ� :
			<input type="text" class="text" name="itembarcode" value="<%= itembarcode %>" size="20" maxlength="32">
		</td>

		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="submitFrm()">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
		    (����)����� : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		    &nbsp;&nbsp;
			����:
			<input type="radio" name="errType" value="" <%= Chkiif(errType = "","checked","") %> > ��ü
			<input type="radio" name="errType" value="S" <%= Chkiif(errType = "S","checked","") %> > ����
        	<!-- input type="radio" name="errType" value="B" <%= Chkiif(errType = "B","checked","") %> > �ҷ� -->
        	<input type="radio" name="errType" value="D" <%= Chkiif(errType = "D","checked","") %> > ����
		</td>
	</tr>
	</form>
</table>
<!-- �˻� �� -->
<% if C_ADMIN_AUTH then %>
<table class="a">
<form name="frm1" method="get" onsubmit="return false;">
<tr>
<td>
[�����ں�] �󿡷��α��Է�
<input type="text" name="bufYYYYMMDD" value="" size="10" maxlength="10">
<input type="button" value="�Է�" onClick="inputDummiErr(this,'<%= shopid %>', '<%= itemgubun %>','<%= itemid %>','<%= itemoption %>')">
&nbsp;
&nbsp; <a href="#" onClick="document.frm1.bufYYYYMMDD.value='2013-01-31'">2013-01-31</a>
&nbsp; <a href="#" onClick="document.frm1.bufYYYYMMDD.value='2013-02-28'">2013-02-28</a>
&nbsp; <a href="#" onClick="document.frm1.bufYYYYMMDD.value='2013-03-31'">2013-03-31</a>
&nbsp; <a href="#" onClick="document.frm1.bufYYYYMMDD.value='2013-04-30'">2013-04-30</a>
&nbsp; <a href="#" onClick="document.frm1.bufYYYYMMDD.value='2013-05-31'">2013-05-31</a>
&nbsp; <a href="#" onClick="document.frm1.bufYYYYMMDD.value='2013-06-30'">2013-06-30</a>
&nbsp; <a href="#" onClick="document.frm1.bufYYYYMMDD.value='2013-07-31'">2013-07-31</a>
&nbsp; <a href="#" onClick="document.frm1.bufYYYYMMDD.value='2013-08-31'">2013-08-31</a>
&nbsp; <a href="#" onClick="document.frm1.bufYYYYMMDD.value='2013-09-30'">2013-09-30</a>
</td>
</tr>
</form>
<% if (oOffStock.FResultCount=1) then %>
<tr>
<form name="frm2" method="get" onsubmit="return false;">
<td>
[�����ں�] ������¥����
<input type="text" name="bufYYYYMMDD" value="2013-02-28" size="10" maxlength="10">
<input type="text" name="bufNo" value="<%=oOffStock.FItemList(i).Ferrrealcheckno%>" size="4" maxlength="9">
<input type="button" value="�Է�" onClick="inputDummiWithLstErrNo(this,'<%= shopid %>', '<%= itemgubun %>','<%= itemid %>','<%= itemoption %>','<%= oOffStock.FItemList(i).Fyyyymmdd %>')">
&nbsp;
&nbsp; <a href="#" onClick="document.frm2.bufYYYYMMDD.value='2013-01-31'">2013-01-31</a>
&nbsp; <a href="#" onClick="document.frm2.bufYYYYMMDD.value='2013-02-28'">2013-02-28</a>
&nbsp; <a href="#" onClick="document.frm2.bufYYYYMMDD.value='2013-03-31'">2013-03-31</a>
&nbsp; <a href="#" onClick="document.frm2.bufYYYYMMDD.value='2013-04-30'">2013-04-30</a>
&nbsp; <a href="#" onClick="document.frm2.bufYYYYMMDD.value='2013-05-31'">2013-05-31</a>
&nbsp; <a href="#" onClick="document.frm2.bufYYYYMMDD.value='2013-06-30'">2013-06-30</a>
&nbsp; <a href="#" onClick="document.frm2.bufYYYYMMDD.value='2013-07-31'">2013-07-31</a>
&nbsp; <a href="#" onClick="document.frm2.bufYYYYMMDD.value='2013-08-31'">2013-08-31</a>
&nbsp; <a href="#" onClick="document.frm2.bufYYYYMMDD.value='2013-09-30'">2013-09-30</a>
</td>
</tr>
</form>
<% end if %>

</table>
<% end if %>

<p>

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frmList">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="65">�����</td>
		<td width="100">����ID</td>
		<td width="100">�귣��ID</td>
		<td width="80">�ŷ�<br>����</td>
		<td width="90">��ǰ<br>�ڵ�</td>
		<td>�����۸�</td>
		<td>�ɼ�</td>
		<td width="30">����</td>
		<!--td width="30">�ҷ�</td-->
		<td width="30">����</td>
		<td width="80">���ID/<br>����ID</td>
		<td width="60">����</td>
		<td width="60">����</td>
		<!--
		<td width="50">����</td>
		<td width="50">����</td>
		-->
    </tr>
	<% for i=0 to oOffStock.FResultCount - 1 %>
	<%
	TotErrsampleitemno   = TotErrsampleitemno + oOffStock.FItemList(i).Ferrsampleitemno
	TotErrbaditemno   = TotErrbaditemno + oOffStock.FItemList(i).Ferrbaditemno
	TotErrrealcheckno = TotErrrealcheckno + oOffStock.FItemList(i).Ferrrealcheckno

	%>
    <tr align="center" bgcolor="#FFFFFF">
		<td><%= oOffStock.FItemList(i).Fyyyymmdd %></td>
		<td><%= oOffStock.FItemList(i).Fshopid %></td>
		<td><%= oOffStock.FItemList(i).Fmakerid %></td>
		<td><%= oOffStock.FItemList(i).fcomm_name %></td>
		<td>
			<%= oOffStock.FItemList(i).fitemgubun %><%= CHKIIF(oOffStock.FItemList(i).fitemid>=1000000,format00(8,oOffStock.FItemList(i).fitemid),format00(6,oOffStock.FItemList(i).fitemid)) %><%= oOffStock.FItemList(i).fitemoption %>
		</td>
		<td align="left">
			<a href="javascript:popShopCurrentStock('<%= oOffStock.FItemList(i).FShopid %>','<%= oOffStock.FItemList(i).Fitemgubun %>','<%= oOffStock.FItemList(i).FItemID %>','<%= oOffStock.FItemList(i).FItemOption %>');">
			<%= oOffStock.FItemList(i).FShopItemname %></a>
		</td>
		<td><%= oOffStock.FItemList(i).FShopItemOptionName %></td>
		<td><%= oOffStock.FItemList(i).Ferrsampleitemno %></td>
		<!-- td><%= oOffStock.FItemList(i).Ferrbaditemno %></td -->
		<td><%= oOffStock.FItemList(i).Ferrrealcheckno %></td>
		<td><%= oOffStock.FItemList(i).FRegUserID %><br><%= oOffStock.FItemList(i).FModiUserID %></td>
        <td>
            <% if C_ADMIN_AUTH then %>
            <input type="text" class="text" name="errsampleitemno" value="<%= oOffStock.FItemList(i).Ferrsampleitemno %>" size="3" >
            <input type="button" class="button" value="����" onClick="EditErrDetail('<%= oOffStock.FItemList(i).Fyyyymmdd %>','<%= oOffStock.FItemList(i).FItemgubun %>','<%= oOffStock.FItemList(i).FItemid %>','<%= oOffStock.FItemList(i).FItemoption %>','<%= oOffStock.FItemList(i).FShopid %>',<%= i %>);">
            <% end if %>
        </td>

        <td>
            <% if C_ADMIN_AUTH then %>
            <input type="text" class="text" name="errrealcheckno" value="<%= oOffStock.FItemList(i).Ferrrealcheckno %>" size="3" >
            <input type="button" class="button" value="����" onClick="EditErrDetail('<%= oOffStock.FItemList(i).Fyyyymmdd %>','<%= oOffStock.FItemList(i).FItemgubun %>','<%= oOffStock.FItemList(i).FItemid %>','<%= oOffStock.FItemList(i).FItemoption %>','<%= oOffStock.FItemList(i).FShopid %>',<%= i %>);">
            <% end if %>
        </td>

		<!--
		<td></td>
      	<td><a href="javascript:DelErrDetail('<%= oOffStock.FItemList(i).Fyyyymmdd %>','<%= oOffStock.FItemList(i).FItemgubun %>','<%= oOffStock.FItemList(i).FItemid %>','<%= oOffStock.FItemList(i).FItemoption %>','<%= oOffStock.FItemList(i).FShopid %>');"><img src="/images/icon_delete.gif" width="45" border="0"></a></td>
      	-->
    </tr>
   	<% next %>
	<tr align="center" bgcolor="#FFFFFF">
	  <td>Total</td>
	  <td colspan="6"></td>
	  <td><%= TotErrsampleitemno %></td>
	  <!--td><%= TotErrbaditemno %></td-->
	  <td><%= TotErrrealcheckno %></td>
	  <td></td>
	  <td></td>
	  <td></td>
	  <!--

	  <td></td>
	  <td></td>
	  -->
	</tr>
</form>
</table>
<form name="frmDel" method="post" action="stockrefresh_process.asp">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="yyyymmdd" value="">
	<input type="hidden" name="shopid" value="">
	<input type="hidden" name="itemgubun" value="">
	<input type="hidden" name="itemid" value="">
	<input type="hidden" name="itemoption" value="">
	<input type="hidden" name="errrealcheckno" value="">
	<input type="hidden" name="errsampleitemno" value="">
	<input type="hidden" name="preyyyymmdd" value="">
</form>

<%
set oOffStock = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
