<%@ language=vbscript %>
<% option explicit %>
<%
'####################################################
' Description :  �������� ���� �ֹ��� �� ���������� NO ����¡ ����
' History : 2009.04.07 ������ ����
'			2010.03.26 �ѿ�� ����
'####################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/offshopsellcls.asp"-->
<!-- #include virtual="/lib/classes/maechul/pgdatacls.asp"-->
<%
dim shopid, oldlist , datefg , prejumunno , makerid , menupos ,yyyy1,mm1,dd1,yyyy2,mm2,dd2, toDate,fromDate
dim cardMinusTotal, cashMinusTotal, cardMinusCnt, cashMinusCnt, buyergubun
dim etcTotal, etcCnt, etcMinusTotal, etcMinusCnt ,i,totalsum ,cardtotal, cashtotal, cardcnt, cashcnt
dim cardpayonly, excmatchfinish, logidx, showdetail, cardsum
dim research
	yyyy1 = request("yyyy1")
	mm1 = request("mm1")
	dd1 = request("dd1")
	yyyy2 = request("yyyy2")
	mm2 = request("mm2")
	dd2 = request("dd2")
	menupos = request("menupos")
	shopid = request("shopid")
	oldlist = request("oldlist")
	datefg = request("datefg")
	makerid = request("makerid")
	buyergubun = request("buyergubun")

	cardpayonly = request("cardpayonly")
	excmatchfinish = request("excmatchfinish")
	logidx = request("logidx")
	showdetail = request("showdetail")

	cardsum = request("cardsum")

	if (cardsum <> "") then
		cardsum = Trim(Replace(cardsum, ",", ""))

		if Not IsNumeric(cardsum) then
			response.write "<script>alert('�ݾ��� ���ڸ� �����մϴ�.');</script>"
			cardsum = ""
		end if
	end if

	research = request("research")

if (research = "") then
	cardpayonly = "Y"
	excmatchfinish = "Y"
end if

if datefg = "" then datefg = "maechul"
''if datefg = "" then datefg = "jumun"

if (yyyy1="") then
	fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())), Cstr(day(now()))-7)
else
	fromDate = DateSerial(yyyy1, mm1, dd1)
end if

if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

toDate = DateSerial(yyyy2, mm2, dd2+1)

yyyy1 = left(fromDate,4)
mm1 = Mid(fromDate,6,2)
dd1 = Mid(fromDate,9,2)

'/����
if (C_IS_SHOP) then

	'//�������϶�
	if C_IS_OWN_SHOP then

		'/���α��� ���� �̸�
		'if getlevel_sn("",session("ssBctId")) > 6 then
			'shopid = C_STREETSHOPID		'"streetshop011"
		'end if
	else
		shopid = C_STREETSHOPID
	end if
else
	'/��ü
	if (C_IS_Maker_Upche) then
		makerid = session("ssBctID")
	else
		if (Not C_ADMIN_USER) then
		    shopid = "X"                ''�ٸ�������ȸ ����.
		else
		end if
	end if
end if

''��Ÿ�� ������ȸ ����
Dim isFixShopView
IF (session("ssBctID")="doota01") then
    shopid="streetshop014"
    C_IS_SHOP = TRUE
    isFixShopView = TRUE
ENd If

dim ooffsell
set ooffsell = new COffShopSellReport
	ooffsell.FRectShopid = shopid
	ooffsell.FRectNormalOnly = "on"
	ooffsell.FRectOldData = oldlist
    ooffsell.FRectTerms = ""
    ooffsell.FRectStartDay = fromDate
    ooffsell.frectdatefg = datefg
    ooffsell.FRectDesigner = makerid
    ooffsell.FRectEndDay = toDate
    ooffsell.FRectbuyergubun = buyergubun

	ooffsell.FRectCardPayOnly = cardpayonly
	ooffsell.FRectExcMatchFinish = excmatchfinish

	''ooffsell.FRectCardSum = cardsum
	ooffsell.FRectPaySum = cardsum
    ooffsell.FRectPgDataCheck ="on"  ''�������߰�

	ooffsell.GetDaylySellJumunList

totalsum =0
cardtotal =0
cashtotal =0
cardcnt   =0
cashcnt   =0
cardMinusTotal =0
cashMinusTotal =0
cardMinusCnt   =0
cashMinusCnt   =0
etcTotal        =0
etcCnt          =0
etcMinusTotal   =0
etcMinusCnt     =0

''response.write logidx & "aaa"

Dim oCPGData
set oCPGData = new CPGData

	oCPGData.FRectIdx = logidx

	if (logidx <> "") then
    	oCPGData.getPGDataOne_OFF
	end if

%>

<script language="javascript">

function frmsubmit(){

	frm.submit();
}

function jsMatchThis(orderno, cardsum) {
	var frm = document.frmAct;
	<% if (oCPGData.FResultCount > 0) then %>
	var pgcardsum = "<%= oCPGData.FOneItem.FcardPrice %>";
	<% else %>
	var pgcardsum = "-1";
	<% end if %>
alert(cardsum + '/' + pgcardsum);
	if (cardsum*1 != pgcardsum*1) {
		alert("��Ī�Ұ�!!\n\n�����װ� ���ξ��� ���� �ٸ��ϴ�.");
		return;
	}

	frm.orderno.value = orderno;

	if ((frm.orderno.value == "") || (frm.logidx.value == "")) {
		alert("�߸��� �����Դϴ�.");
		return;
	}

	if (confirm("��Ī�Ͻðڽ��ϱ�?") == true) {
		frm.submit();
	}
}

</script>

<!-- ǥ ��ܹ� ����-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>" class="a">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="logidx" value="<%= logidx %>">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="50" bgcolor="<%= adminColor("gray") %>" rowspan="3">�˻�<br>����</td>
	<td align="left">
		<table border="0" width="100%" cellpadding="3" cellspacing="0" class="a">
		<tr>
			<td>
				* �Ⱓ :
				<% drawmaechuldatefg "datefg" ,datefg ,""%>
				<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
				&nbsp;&nbsp;
				<%
				'����/������
				if (C_IS_SHOP) then
				%>
					<% if (not C_IS_OWN_SHOP and shopid <> "") or (isFixShopView) then %>
						* ���� : <%=shopid%><input type="hidden" name="shopid" value="<%= shopid %>">
					<% else %>
						* ���� : <% drawSelectBoxOffShop "shopid",shopid %>
					<% end if %>
				<% else %>
					* ���� : <% drawSelectBoxOffShop "shopid",shopid %>
				<% end if %>
			</td>
		</tr>
		</table>
    </td>
		<td  width="50" bgcolor="<%= adminColor("gray") %>" rowspan="3">
		<input type="button" class="button_s" value="�˻�" onClick="frmsubmit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<table border="0" width="100%" cellpadding="3" cellspacing="0" class="a">
		<tr>
			<td>
				<% if C_IS_Maker_Upche then %>
					* �귣�� : <%= makerid %><input type="hidden" name="makerid" value="<%= makerid %>">
				<% else %>
					* �귣�� : <% drawSelectBoxDesignerwithName "makerid",makerid %>
				<% end if %>
				&nbsp;&nbsp;
				* ��������: <% drawoffshop_commoncode "buyergubun", buyergubun, "buyergubun", "MAIN", "", " onchange='frmsubmit();'" %>
				&nbsp;&nbsp;
				* ������(ī�� or ����):
				<input type="text" class="text" name="cardsum" value="<%= cardsum %>">
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">
		<table border="0" width="100%" cellpadding="3" cellspacing="0" class="a">
		<tr>
			<td>
				<input type="checkbox" name="excmatchfinish" value="Y" <% if (excmatchfinish = "Y") then %>checked<% end if %> > PG�� ���γ��� ��Ī�Ϸ� ����
				&nbsp;
				<input type="checkbox" name="cardpayonly" value="Y" <% if (cardpayonly = "Y") then %>checked<% end if %> > ī����� ������
				&nbsp;
				<input type="checkbox" name="showdetail" value="Y" <% if (showdetail = "Y") then %>checked<% end if %> > �󼼳��� ǥ��
			</td>
		</tr>
		</table>
	</td>
</tr>
</form>
</table>
<!-- ǥ ��ܹ� ��-->

<br>
<!-- ǥ �߰��� ����-->
<table width="100%" align="center" cellpadding="1" cellspacing="1" class="a">
<tr valign="bottom">
    <td align="left">
		<% if (oCPGData.FResultCount > 0) then %>
			PG�� : <%= oCPGData.FOneItem.FPGgubun %>
			&nbsp;
			PG��KEY : <%= oCPGData.FOneItem.FPGkey %>
			&nbsp;
			�ŷ����� : <%= oCPGData.FOneItem.FappDate %>
			&nbsp;
			�ŷ��� : <b><%= FormatNumber(oCPGData.FOneItem.FcardPrice, 0) %></b>
		<% end if %>

    </td>
    <td align="right">
    </td>
</tr>
</table>

<p><br>

<!-- ǥ �߰��� ��-->
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="#FFFFFF" height="25">
	<td colspan="20">
		�˻���� : <b><%=ooffsell.FresultCount%></b>
	</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td height="25" width="110">
		�����
	</td>
	<td width="110">
		<% if datefg = "maechul" then %>
			<%= chkIIF(shopid="cafe002","������","�ֹ���ȣ") %>
		<% else %>
			<%= chkIIF(shopid="cafe002","�ֹ���","�ֹ���ȣ") %>
		<% end if %>
	</td>
	<td></td>

	<td width="90">���ϸ���</td>
	<td width="90">����Ʈī��</td>
	<td width="90">��ǰ��</td>
	<td width="110">�ſ�ī��</td>
	<td width="110">����</td>

	<td width="70"></td>

	<td rowspan="2" width="70">�ǸŰ�</td>
	<td rowspan="2" width="70">�����</td>
	<% if shopid<>"cafe002" then %>
		<td width="150">�ֹ��Ͻ�</td>
	<% end if %>
	<td>KICC��Ī</td>
    <td rowspan="2">���ι�ȣ</td>
	<td rowspan="2">���<br>(�����ֹ���ȣ)</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td height="25"></td>
	<td height="25"></td>
	<td>�귣��</td>

	<td colspan="5">��ǰ��</td>

	<td>����</td>

	<% if shopid<>"cafe002" then %>
		<td></td>
	<% end if %>
	<td></td>
</tr>
<%
if ooffsell.FresultCount > 0 then

for i=0 to ooffsell.FresultCount-1

if prejumunno<>ooffsell.FItemList(i).ForderNo then

	totalsum = totalsum + ooffsell.FItemList(i).Frealsum
	if (ooffsell.FItemList(i).Fcardsum>0) then
        cardtotal = cardtotal + ooffsell.FItemList(i).Fcardsum
        cardcnt   = cardcnt + 1
    elseif (ooffsell.FItemList(i).Fcardsum<0) then
        cardMinusTotal = cardMinusTotal + ooffsell.FItemList(i).Fcardsum
        cardMinusCnt   =cardMinusCnt + 1
    end if

    if (ooffsell.FItemList(i).Fcashsum>0) then
        cashtotal = cashtotal + ooffsell.FItemList(i).Fcashsum
        cashcnt   = cashcnt + 1
    elseif (ooffsell.FItemList(i).Fcashsum<0) then
        cashMinusTotal = cashMinusTotal + ooffsell.FItemList(i).Fcashsum
        cashMinusCnt   =cashMinusCnt + 1
    end if

    if (ooffsell.FItemList(i).FgiftcardPaysum>0) then
        etcTotal = etcTotal + ooffsell.FItemList(i).FgiftcardPaysum
        etcCnt   = etcCnt + 1
    elseif (ooffsell.FItemList(i).FgiftcardPaysum<0) then
        etcMinusTotal = etcMinusTotal + ooffsell.FItemList(i).FgiftcardPaysum
        etcMinusCnt   =etcMinusCnt + 1
    end if

	prejumunno = ooffsell.FItemList(i).ForderNo

	if IsNull(ooffsell.FItemList(i).Fpointuserno) then
		ooffsell.FItemList(i).Fpointuserno = 0
	end if

%>
<tr align="center" bgcolor="<% if (showdetail = "") then %>FFFFFF<% else %>EEEEEE<% end if %>">
	<td height="25">
		<%= ooffsell.FItemList(i).Fshopid %>
	</td>
	<td>
		<%= chkIIF(shopid="cafe002",ooffsell.FItemList(i).Fshopregdate,ooffsell.FItemList(i).ForderNo) %>
	</td>
	<td><font color="<%= ooffsell.FItemList(i).JumunMethodColor %>"><%= ooffsell.FItemList(i).JumunMethodName %></font></td>

	<td align="right"><%= FormatNumber(ooffsell.FItemList(i).Fspendmile,0) %></td>
	<td align="right"><%= FormatNumber(ooffsell.FItemList(i).FgiftcardPaysum,0) %></td>
	<td align="right"><%= FormatNumber(ooffsell.FItemList(i).FTenGiftCardPaySum,0) %></td>
	<td align="right"><b><%= FormatNumber(ooffsell.FItemList(i).Fcardsum,0) %></b></td>
	<td align="right"><%= FormatNumber(ooffsell.FItemList(i).Fcashsum,0) %></td>

	<td><input type="button" class="button" value="��Ī" onClick="jsMatchThis('<%= ooffsell.FItemList(i).ForderNo %>', '<%= ooffsell.FItemList(i).Fcardsum %>')" <% if (ooffsell.FItemList(i).FmatchCount > 0) then %>disabled<% end if %> ></td>

	<td align="right"><%= FormatNumber(ooffsell.FItemList(i).Ftotalsum,0) %></td>
	<td align="right"><b><%= FormatNumber(ooffsell.FItemList(i).Frealsum,0) %></b></td>
	<% if shopid<>"cafe002" then %>
		<td><%= ooffsell.FItemList(i).Fshopregdate %></td>
	<% end if %>
	<td>
		<% if (ooffsell.FItemList(i).FmatchCount > 0) then %>Y<% end if %>
	</td>

    <td><%=ooffsell.FItemList(i).Fcardappno%></td>
	<td><%=ooffsell.FItemList(i).Freforderno%></td>
</tr>
<% end if %>
<% if (showdetail = "Y") then %>
<tr align="center" bgcolor="FFFFFF">
	<td height="25"></td>
	<td height="25"></td>
	<td><%= ooffsell.FItemList(i).FMakerID %></td>

	<td colspan="5" align="left">&nbsp; <%= ooffsell.FItemList(i).FItemName %> <%= ooffsell.FItemList(i).FItemOptionName %></td>

	<% if ooffsell.FItemList(i).FItemNo<0 then %>
		<td align="center"><font color=red><%= ooffsell.FItemList(i).FItemNo %></font></td>
		<td align="right"><font color=red><%= FormatNumber(ooffsell.FItemList(i).FSellPrice,0) %></font></td>
		<td align="right"><font color=red><%= FormatNumber(ooffsell.FItemList(i).FRealSellPrice,0) %></font></td>
	<% else %>
		<td align="center"><%= ooffsell.FItemList(i).FItemNo %></td>
		<td align="right"><%= FormatNumber(ooffsell.FItemList(i).FSellPrice,0) %></td>
		<td align="right"><%= FormatNumber(ooffsell.FItemList(i).FRealSellPrice,0) %></td>
	<% end if %>

	<% if shopid<>"cafe002" then %>
	<td colspan="4"></td>
	<% else %>
	<td colspan="3"></td>
	<% end if %>
</tr>
<% end if %>
<% next %>
<tr bgcolor="#FFFFFF" align="center">
	<td colspan="3"><b>�Ѱ�</b></td>
	<td colspan="12" align="right">
		<table width=440 border=0 cellspacing=0 cellpadding=0 class="a">
		<tr>
		    <td>���� :</td>
		    <td align="right"><%= FormatNumber(cashtotal,0) %> ��</td>
		    <td align="center">(<%= FormatNumber(cashcnt,0) %> ��)</td>
		    <td width=10></td>
		    <td align="right"><font color="red"><%= FormatNumber(cashMinusTotal,0) %> ��</td>
		    <td align="center"><font color="red">(<%= FormatNumber(cashMinusCnt,0) %> ��)</font></td>
		    <td align="right"><%= FormatNumber(cashtotal + cashMinusTotal,0) %> ��</td>
		</tr>
		<tr>
		    <td>ī�� :</td>
		    <td align="right"><%= FormatNumber(cardtotal,0) %> ��</td>
		    <td align="center">(<%= FormatNumber(cardcnt,0) %> ��)</td>
		    <td></td>
		    <td align="right"><font color="red"><%= FormatNumber(cardMinusTotal,0) %> ��</font></td>
		    <td align="center"><font color="red">(<%= FormatNumber(cardMinusCnt,0) %> ��)</font></td>
		    <td align="right"><%= FormatNumber(cardtotal + cardMinusTotal,0) %> ��</td>
		</tr>
		<tr>
		    <td>��ǰ�� :</td>
		    <td align="right"><%= FormatNumber(etcTotal,0) %> ��</td>
		    <td align="center">(<%= FormatNumber(etccnt,0) %> ��)</td>
		    <td></td>
		    <td align="right"><font color="red"><%= FormatNumber(etcMinusTotal,0) %> ��</font></td>
		    <td align="center"><font color="red">(<%= FormatNumber(etcMinusCnt,0) %> ��)</font></td>
		    <td align="right"><%= FormatNumber(etcTotal + etcMinusTotal,0) %> ��</td>
		</tr>
		<tr>
		    <td>�հ� :</td>
		    <td align="right"><%= FormatNumber(cashtotal + cardtotal + etcTotal,0) %> ��</td>
		    <td align="center">(<%= FormatNumber(cashcnt + cardcnt + etccnt,0) %> ��)</td>
		    <td></td>
		    <td align="right"><font color="red"><%= FormatNumber(cashMinusTotal + cardMinusTotal + etcMinusTotal,0) %> ��</td>
		    <td align="center"><font color="red">(<%= FormatNumber(cashMinusCnt + cardMinusCnt + etcMinusCnt,0) %> ��)</font></td>
		    <td align="right"><%= FormatNumber(totalsum,0) %> ��</td>
		</tr>
		</table>
	</td>
</tr>
<% else %>
<tr  align="center" bgcolor="#FFFFFF">
	<td colspan="20">��ϵ� ������ �����ϴ�.</td>
</tr>
<% end if %>
</table>

<form name="frmAct" method="post" action="<%=stsAdmURL%>/admin/maechul/pgdata/pgdata_process.asp">
<input type="hidden" name="mode" value="matchoneorder">
<input type="hidden" name="logidx" value="<%= logidx %>">
<input type="hidden" name="orderno" value="">
</form>

<%
set ooffsell = Nothing
%>
<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
