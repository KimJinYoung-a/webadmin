<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionSTAdmin.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/maechul/pgdatacls.asp"-->
<%

dim page, research, i
dim yyyy1, mm1
'', stplace, targetGbn, itemgubun
''dim ipgoMWdiv, itemMWdiv, itemid
''dim startYYYYMMDD, endYYYYMMDD
''dim addInfoType
''dim lastmwdiv, lastmakerid
dim tmpDate, nextMonth, prevMonth


page       	= requestCheckvar(request("page"),10)
research	= requestCheckvar(request("research"),10)
yyyy1       = requestCheckvar(request("yyyy1"),10)
mm1         = requestCheckvar(request("mm1"),10)


if (page="") then page = 1
if (yyyy1="") then
	tmpDate = Left(DateAdd("m", 0, Now()), 7)
	yyyy1 = Left(tmpDate, 4)
	mm1 = Right(tmpDate, 2)
end if

'// ============================================================================
dim opgdataAdvPrice
set opgdataAdvPrice = new CPGData

opgdataAdvPrice.FPageSize = 100
opgdataAdvPrice.FCurrPage = 1
opgdataAdvPrice.FRectYYYYMM = yyyy1 + "-" + mm1

opgdataAdvPrice.getPGDataAdvPriceList

dim fromDate, toDate, showDiffPopup

fromDate = DateSerial(yyyy1, mm1, 1)
toDate = DateAdd("d", -1, DateAdd("m", 1, fromDate))

prevMonth = DateAdd("m", -1, fromDate)
nextMonth = DateAdd("m", 1, fromDate)

%>

<script language='javascript'>

function jsPopCheckPayLog(pggubun) {
    var pop;

    pop = window.open("/admin/maechul/payment_maechul_log_chk.asp?menupos=4161&pggubun=" + pggubun + "&yyyy1=<%= Year(fromDate) %>&mm1=<%= Month(fromDate) %>&dd1=<%= Day(fromDate) %>&yyyy2=<%= Year(toDate) %>&mm2=<%= Month(toDate) %>&dd2=<%= Day(toDate) %>")
}

function jsMakeAdvPrice(v) {
	var frm = document.getElementById('frmAct');
	var i;

	if (frm == undefined) {
		alert('============================== \n\n�� �� ���� �����Դϴ�.\n\n ==============================');
		return;
	}

	frm.mode.value = "makeadvprc" + v;

	if (confirm("�ۼ��Ͻðڽ��ϱ�?") == true) {
		frm.submit();
	}
}

function popAppPriceDetail(yyyymm, targetGbn, pggubun, pguserid) {
	var yyyy, mm;
	var srcGbn;
	var lastDayOfMonth

	yyyy = yyyymm.substring(0, 4);
	mm = yyyymm.substring(5);

	var tmpDate = new Date(yyyy, mm, 0);
	lastDayOfMonth = tmpDate.getDate();

	switch (pguserid) {
		case "balance":
		case "giftcard":
		case "mileage":
			// ��������Ʈ��ġ�ݰ���
			switch (pguserid) {
				case "balance":
					srcGbn = "D";
					break;
				case "giftcard":
					srcGbn = "G";
					break;
				default:
					srcGbn = "M";
			}

			window.open("/admin/maechul/managementsupport/combine_point_deposit_month.asp?menupos=1612&yyyy1=" + yyyy + "&mm1=" + mm + "&yyyy2=" + yyyy + "&mm2=" + mm + "&srcGbn=" + srcGbn + "&targetGbn=" + targetGbn,"popAppPriceDetail","width=1400,height=580,scrollbars=yes,resizabled=yes");
			break;
		case "bankipkum_10x10":
		case "bankrefund_10x10":
		case "bankipkum_fingers":
		case "bankrefund_fingers":
		case "gifticon":
		case "giftting":
		case "inicis":
		case "okcashbag":
		case "uplus":
		case "teenxteen3":
		case "teenxteen4":
		case "teenxteen6":
		case "teenxteen8":
		case "teenxteen9":
		case "tenbyten01":
		case "tenbyten02":
		case "R5523":
		case "KB":
		case "NH":
		case "LOTTE":
		case "BC":
		case "SAMSUNG":
		case "SHINHAN":
		case "KE":
		case "HANA":
		case "HYUNDAI":
			// �¶���/�������� ���γ���
			if (targetGbn == "OF") {
				window.open("/admin/maechul/pgdata/pgdata_statistics_off.asp?menupos=1565&page=&research=on&dategubun=appdate&chkSearchAppDate=Y&yyyy1=" + yyyy + "&mm1=" + mm + "&dd1=01&yyyy2=" + yyyy + "&mm2=" + mm + "&dd2=" + lastDayOfMonth + "&yyyy3=" + yyyy + "&mm3=" + mm + "&dd3=01&yyyy4=" + yyyy + "&mm4=" + mm + "&dd4=" + lastDayOfMonth,"popAppPriceDetail","width=1400,height=580,scrollbars=yes,resizabled=yes");
			} else {
				window.open("/admin/maechul/pgdata/pgdata_statistics_on.asp?menupos=1572&sumgubun=appMethod&dategubun=appdate&chkSearchAppDate=Y&yyyy1=" + yyyy + "&mm1=" + mm + "&dd1=01&yyyy2=" + yyyy + "&mm2=" + mm + "&dd2=" + lastDayOfMonth + "&yyyy3=" + yyyy + "&mm3=" + mm + "&dd3=01&yyyy4=" + yyyy + "&mm4=" + mm + "&dd4=" + lastDayOfMonth + "&pggubun=" + pggubun + "&PGuserid=" + pguserid,"popAppPriceDetail","width=1400,height=580,scrollbars=yes,resizabled=yes");
			}
			break;
		case "CASH":
		case "happymoney":
		case "streetshop018":
		case "partner":
			window.open("/common/offshop/maechul/statistic/statistic_checkmethod_datamart.asp?reload=on&menupos=1541&datefg=maechul&yyyy1=" + yyyy + "&mm1=" + mm + "&dd1=01&yyyy2=" + yyyy + "&mm2=" + mm + "&dd2=" + lastDayOfMonth + "&shopid=&offgubun=1&BanPum=&inc3pl=","popAppPriceDetail","width=1400,height=580,scrollbars=yes,resizabled=yes");
			break;
		default:
			alert("ǥ���� ��������");
			break;
	}
}

function popAppPriceDetailSUM(yyyymm, targetGbn, pggubun, pguserid) {
	var yyyy, mm;
	var srcGbn;
	var lastDayOfMonth;
	var pop;

	yyyy = yyyymm.substring(0, 4);
	mm = yyyymm.substring(5);

	var tmpDate = new Date(yyyy, mm, 0);
	lastDayOfMonth = tmpDate.getDate();

	if (targetGbn == 'ON') {
		switch (pguserid) {
			case "balance":
			case "giftcard":
			case "mileage":
				// ��������Ʈ��ġ�ݰ���
				switch (pguserid) {
					case "balance":
						srcGbn = "D";
						break;
					case "giftcard":
						srcGbn = "G";
						break;
					default:
						srcGbn = "M";
				}
				pop = window.open("/admin/maechul/managementsupport/combine_point_deposit_month.asp?menupos=1612&yyyy1=" + yyyy + "&mm1=" + mm + "&yyyy2=" + yyyy + "&mm2=" + mm + "&srcGbn=" + srcGbn + "&targetGbn=" + targetGbn,"popAppPriceDetailSUM","width=1400,height=580,scrollbars=yes,resizabled=yes");
				break;
			default:
				pop = window.open("/admin/maechul/pgdata/pgdata_statistics_on.asp?menupos=1572&sumgubun=appMethod&dategubun=appdate&chkSearchAppDate=Y&yyyy1=" + yyyy + "&mm1=" + mm + "&dd1=01&yyyy2=" + yyyy + "&mm2=" + mm + "&dd2=" + lastDayOfMonth + "&yyyy3=" + yyyy + "&mm3=" + mm + "&dd3=01&yyyy4=" + yyyy + "&mm4=" + mm + "&dd4=" + lastDayOfMonth + "&pggubun=" + pggubun + "&PGuserid=" + pguserid,"popAppPriceDetailSUM","width=1400,height=580,scrollbars=yes,resizabled=yes");
				break;
		}
	} else if (targetGbn == 'OF') {
		switch (pguserid) {
			case "balance":
			case "giftcard":
			case "mileage":
				// ��������Ʈ��ġ�ݰ���
				switch (pguserid) {
					case "balance":
						srcGbn = "D";
						break;
					case "giftcard":
						srcGbn = "G";
						break;
					default:
						srcGbn = "M";
				}
				pop = window.open("/admin/maechul/managementsupport/combine_point_deposit_month.asp?menupos=1612&yyyy1=" + yyyy + "&mm1=" + mm + "&yyyy2=" + yyyy + "&mm2=" + mm + "&srcGbn=" + srcGbn + "&targetGbn=" + targetGbn,"popAppPriceDetailSUM","width=1400,height=580,scrollbars=yes,resizabled=yes");
				break;
			default:
				pop = window.open("/admin/maechul/pgdata/pgdata_statistics_off.asp?menupos=1565&page=&research=on&dategubun=appdate&chkSearchAppDate=Y&yyyy1=" + yyyy + "&mm1=" + mm + "&dd1=01&yyyy2=" + yyyy + "&mm2=" + mm + "&dd2=" + lastDayOfMonth + "&yyyy3=" + yyyy + "&mm3=" + mm + "&dd3=01&yyyy4=" + yyyy + "&mm4=" + mm + "&dd4=" + lastDayOfMonth + "&pggubun=" + pggubun + "&PGuserid=" + pguserid,"popAppPriceDetailSUM","width=1400,height=580,scrollbars=yes,resizabled=yes");
				break;
		}
	} else {
		alert("ǥ���� ��������(" + targetGbn + ")");
		return;
	}
	pop.focus();
}

function popMeachulPriceDetail(yyyymm, targetGbn, pggubun, pguserid) {
	var yyyy, mm;
	var lastDayOfMonth

	yyyy = yyyymm.substring(0, 4);
	mm = yyyymm.substring(5) * 1;

	var tmpDate = new Date(yyyy, mm, 0);
	lastDayOfMonth = tmpDate.getDate();

	switch (targetGbn) {
		case "ON":
		case "AC":
			switch (pguserid) {
				case "balance":
				case "bankipkum_10x10":
				case "bankrefund_10x10":
				case "giftcard":
				case "gifticon":
				case "giftting":
				case "teenxteen3":
				case "teenxteen4":
				case "teenxteen6":
				case "teenxteen8":
				case "teenxteen9":
				case "mileage":
				case "okcashbag":
				case "tenbyten01":
				case "tenbyten02":
				case "R5523":
				case "bankrefund_fingers":
				case "":
					window.open("/admin/maechul/maechul_month_paymentPG_log.asp?menupos=1625&selD=2&selSY=" + yyyy + "&selSM=" + mm + "&selEY=" + yyyy + "&selEM=" + mm + "&selPGC=&selPGID=" + pguserid,"popAppPriceDetail","width=1400,height=580,scrollbars=yes,resizable=yes");
					break;
				default:
					alert("ǥ���� ��������");
					break;
			}
			break;
		default:
			alert("ǥ���� ��������");
			break;
	}
}

function popMeachulPriceReasonDetail(yyyymm, targetGbn, pggubun, pguserid, reasonGubun) {
	var yyyy, mm;
	var lastDayOfMonth;
	var pop;

	yyyy = yyyymm.substring(0, 4);
	mm = yyyymm.substring(5) * 1;

	var tmpDate = new Date(yyyy, mm, 0);
	lastDayOfMonth = tmpDate.getDate();

	switch (targetGbn) {
		case "ON":
			pop = window.open("/admin/maechul/pgdata/pgdata_on.asp?menupos=1567&chkSearchAppDate=Y&yyyy1=" + yyyy + "&mm1=" + mm + "&dd1=01&yyyy2=" + yyyy + "&mm2=" + mm + "&dd2=" + lastDayOfMonth + "&yyyy3=" + yyyy + "&mm3=" + mm + "&dd3=01&yyyy4=" + yyyy + "&mm4=" + mm + "&dd4=" + lastDayOfMonth + "&PGuserid=" + pguserid + "&reasonGubun=" + reasonGubun + "" + "&pggubun=" + pggubun,"popAppPriceDetail","width=1400,height=580,scrollbars=yes,resizable=yes");
			break;
		case "AC":
		default:
			pop = window.open("/admin/maechul/pgdata/pgdata_off.asp?menupos=1562&dateType=A&yyyy1=" + yyyy + "&mm1=" + mm + "&dd1=01&yyyy2=" + yyyy + "&mm2=" + mm + "&dd2=" + lastDayOfMonth + "&reasonGubun=" + reasonGubun + "&pggubun=" + pggubun + "&pguserid=" + pguserid,"popAppPriceDetail","width=1400,height=580,scrollbars=yes,resizable=yes");
			break;
	}
	pop.focus();
}

function jsPopShowDiff(yyyymm, targetGbn, PGgubun, pguserid) {
	var yyyy, mm;
	var lastDayOfMonth;
	var pop;

	yyyy = yyyymm.substring(0, 4);
	mm = yyyymm.substring(5) * 1;

	var tmpDate = new Date(yyyy, mm, 0);
	lastDayOfMonth = tmpDate.getDate();

	pop = window.open("/admin/maechul/maechul_payment_log.asp?menupos=1606&dateGubun=payreqdate&matchState=Y&showOnlyPriceNotMatch=Y&targetGbn=" + targetGbn + "&yyyy1=" + yyyy + "&mm1=" + mm + "&dd1=01&yyyy2=" + yyyy + "&mm2=" + mm + "&dd2=" + lastDayOfMonth + "&PGuserid=" + pguserid + "&PGgubun=" + PGgubun,"popAppPriceDetail","width=1400,height=580,scrollbars=yes,resizable=yes");
	pop.focus();
}

function jsGotoPrevMonth() {
    var frm = document.frm;
    var yyyy, mm;

    yyyy = <%= Year(prevMonth) %>;
    mm = <%= Month(prevMonth) %>;

    frm.yyyy1.value = yyyy;
    frm.mm1.value = (mm < 10 ? "0" : "") + mm;

    frm.submit();
}

function jsGotoNextMonth() {
    var frm = document.frm;
    var yyyy, mm;

    yyyy = <%= Year(nextMonth) %>;
    mm = <%= Month(nextMonth) %>;

    frm.yyyy1.value = yyyy;
    frm.mm1.value = (mm < 10 ? "0" : "") + mm;

    frm.submit();
}

</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<input type="hidden" name="research" value="on">
	<input type="hidden" name="page" value="">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="1" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
		<td align="left">
			&nbsp;
			<font color="#CC3333">��/�� :</font> <% DrawYMBox yyyy1,mm1 %> ��
            <input type="button" class="button" value="����" onClick="jsGotoPrevMonth()">
            <input type="button" class="button" value="������" onClick="jsGotoNextMonth()">
		</td>
		<td rowspan="1" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	</form>
</table>
<!-- �˻� �� -->
<p />

* <font color="red">���� ���ݸ����� ����</font>�� ���, ���� �귣�������� �⺻�������� �����Ǿ� �ִ��� Ȯ���ϼ���.<br /><br />

* PG���ξ�(��ġ��, ����Ʈ, ���ϸ���) : ���տ�ġ�ݰ���<br />
* PG���ξ�(�������Ա�, ������ȯ��, �ſ�ī��, īī������ ��) : PG�� ���γ���<br />
* �հ� : ������(����) + ������(���� �̿�), CS���� ��<br />
* ������(����) : �����α� �ǽ��ξ�<br />
* ������(���� �̿�), CS���� �� : PG�� ���γ��� ���� ���� �̿� �Է� ��<br /><br />

* �������� ����Ʈī�� ������ ���� �ִ� ���, �ֹ���ȣ ��Ī �� �����Է� �� �̼������ۼ��ϸ� �����Էµ˴ϴ�.<br /><br />

* ����(��ġ��, ����Ʈ, ���ϸ���) : �����α� �����Ǿ�����, ���տ�ġ�� ���ۼ� �ȵ� ���̽�<br />
* ����(�ſ�ī�� ��) : ���ξ� ū ��� : ���γ��� ������ �����α� ���� ���̽�, �Ǵ� �����α�-���γ��� ��Ī �ȵ� ���̽�<br />
* ����(�ſ�ī�� ��) : ���ξ� ���� ��� : �����α� ������ ���γ���  ���� ���̽�<br />

<p />

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">
			<input type="button" class="button" value="���ۼ� 01(<%= yyyy1 %>-<%= mm1 %>)" onclick="jsMakeAdvPrice('01');">
			<input type="button" class="button" value="���ۼ� 02(<%= yyyy1 %>-<%= mm1 %>)" onclick="jsMakeAdvPrice('02');">
			<input type="button" class="button" value="���ۼ� 03(<%= yyyy1 %>-<%= mm1 %>)" onclick="jsMakeAdvPrice('03');">
		</td>
		<td align="right">

		</td>
	</tr>
</table>
<!-- �׼� �� -->

<p>

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="25">
		<td width="60">���</td>
		<td width="40">����<br />����</td>
		<td width="120">PG��</td>
		<td width="120">PG��ID</td>
		<td width="110">PG���ξ�</td>
		<td width="5"></td>
		<td width=110>����</td>
		<td width="5"></td>
		<td width=110>�հ�</td>
		<td width=110>������<br />(����)<br /><font color="red">(����-����)</font></td><!-- 001 -->
		<td width=110>������<br />(���޻� ����)</td><!-- 002 -->
        <td width=110>������<br />(�̴Ϸ�Ż)</td><!-- 003 -->
		<td width=110>������<br />(��ġ��)</td><!-- 020 -->
		<td width=110>������<br />(��ġ��ȯ��)</td><!-- 025 -->
		<td width=110>������<br />(����Ʈ)</td><!-- 030 -->
		<td width=110>������<br />(����Ʈȯ��)</td><!-- 035 -->
        <td width=110>������<br />(B2B ����)</td><!-- 004 -->
		<td width=110>CS����</td><!-- 040 -->
		<td width=110>���ڼ���</td><!-- 800 -->
		<td width=110>��Ÿ</td><!-- 900 -->
		<td width=110>�ΰŽ�<br />���ݸ���</td><!-- 901 -->
		<td width=110>������<br />��Ȯ��</td><!-- 950 -->
		<td width=110>��Ҹ�Ī</td><!-- 999 -->
		<td width=110>����<br />���Է�</td><!-- XXX -->
		<td width="150">�ۼ���</td>
		<td>���</td>
	</tr>
	<% if opgdataAdvPrice.FResultCount >0 then %>
	<% for i=0 to opgdataAdvPrice.FResultcount-1 %>
    <%
    showDiffPopup = False
    if (opgdataAdvPrice.FItemList(i).FappPrice - opgdataAdvPrice.FItemList(i).GetAdvPriceSUM) <> 0 then
        showDiffPopup = True
    elseif opgdataAdvPrice.FItemList(i).GetDiffIfExist(opgdataAdvPrice.FItemList(i).FpayLogAdvPrice, opgdataAdvPrice.FItemList(i).FpayReqPrice) <> "" then
        showDiffPopup = True
    end if

    %>
	<% if (opgdataAdvPrice.FItemList(i).FtargetGbn = "OF") then %>
	<tr bgcolor="#DDDDFF" height=25>
		<% else %>
		<tr bgcolor="#FFFFFF" height=25>
	<% end if %>
		<td align=center><%= opgdataAdvPrice.FItemList(i).Fyyyymm %></td>
		<td align=center><%= opgdataAdvPrice.FItemList(i).FtargetGbn %></td>
		<td align=center><%= opgdataAdvPrice.FItemList(i).FPGgubun %></td>
		<td align=center><%= opgdataAdvPrice.FItemList(i).FPGuserid %></td>
		<td align=right>
			<a href="javascript:popAppPriceDetailSUM('<%= opgdataAdvPrice.FItemList(i).Fyyyymm %>', '<%= opgdataAdvPrice.FItemList(i).FtargetGbn %>', '<%= opgdataAdvPrice.FItemList(i).FPGgubun %>', '<%= opgdataAdvPrice.FItemList(i).FPGuserid %>');">
				<%= FormatNumber(opgdataAdvPrice.FItemList(i).FappPrice, 0) %>
			</a>
		</td>
		<td align=center></td>
		<td align=right>
			<% if showDiffPopup then %><a href="javascript:jsPopCheckPayLog('<%= opgdataAdvPrice.FItemList(i).FPGgubun %>')"><font color="red"><% end if %>
			<%= FormatNumber((opgdataAdvPrice.FItemList(i).FappPrice - opgdataAdvPrice.FItemList(i).GetAdvPriceSUM), 0) %>
            <% if showDiffPopup then %></font></a><% end if %>
		</td>
		<td align=center></td>
		<td align=right>
			<%= FormatNumber(opgdataAdvPrice.FItemList(i).GetAdvPriceSUM, 0) %>
		</td>
		<td align=right>
			<a href="javascript:popMeachulPriceReasonDetail('<%= opgdataAdvPrice.FItemList(i).Fyyyymm %>', '<%= opgdataAdvPrice.FItemList(i).FtargetGbn %>', '<%= opgdataAdvPrice.FItemList(i).FPGgubun %>', '<%= opgdataAdvPrice.FItemList(i).FPGuserid %>', '001');">
				<%= FormatNumber(opgdataAdvPrice.FItemList(i).GetMeachulPrice(opgdataAdvPrice.FItemList(i).FPGgubun, opgdataAdvPrice.FItemList(i).FpayLogAdvPrice, opgdataAdvPrice.FItemList(i).FreasonGubun001), 0) %>
			</a>
			<% Call opgdataAdvPrice.FItemList(i).ShowDiffIfExistWithPGgubun(opgdataAdvPrice.FItemList(i).FPGgubun, opgdataAdvPrice.FItemList(i).FpayLogAdvPrice, opgdataAdvPrice.FItemList(i).FreasonGubun001) %>
			<% if showDiffPopup then %>
			<a href="javascript:jsPopShowDiff('<%= opgdataAdvPrice.FItemList(i).Fyyyymm %>', '<%= opgdataAdvPrice.FItemList(i).FtargetGbn %>', '<%= opgdataAdvPrice.FItemList(i).FPGgubun %>', '<%= opgdataAdvPrice.FItemList(i).FPGuserid %>')">
				<%= opgdataAdvPrice.FItemList(i).GetDiffIfExist(opgdataAdvPrice.FItemList(i).FpayLogAdvPrice, opgdataAdvPrice.FItemList(i).FpayReqPrice) %>
                <%= CHKIIF(opgdataAdvPrice.FItemList(i).GetDiffIfExist(opgdataAdvPrice.FItemList(i).FpayLogAdvPrice, opgdataAdvPrice.FItemList(i).FpayReqPrice)="", "<br />(ǥ��)", "") %>
			</a>
			<% end if %>
		</td>
		<td align=right>
			<% if Not IsNull(opgdataAdvPrice.FItemList(i).FreasonGubun002) then %>
			<a href="javascript:popMeachulPriceReasonDetail('<%= opgdataAdvPrice.FItemList(i).Fyyyymm %>', '<%= opgdataAdvPrice.FItemList(i).FtargetGbn %>', '<%= opgdataAdvPrice.FItemList(i).FPGgubun %>', '<%= opgdataAdvPrice.FItemList(i).FPGuserid %>', '002');">
				<%= FormatNumber(opgdataAdvPrice.FItemList(i).FreasonGubun002, 0) %>
			</a>
			<% end if %>
		</td>
		<td align=right>
			<% if Not IsNull(opgdataAdvPrice.FItemList(i).FreasonGubun003) then %>
			<a href="javascript:popMeachulPriceReasonDetail('<%= opgdataAdvPrice.FItemList(i).Fyyyymm %>', '<%= opgdataAdvPrice.FItemList(i).FtargetGbn %>', '<%= opgdataAdvPrice.FItemList(i).FPGgubun %>', '<%= opgdataAdvPrice.FItemList(i).FPGuserid %>', '003');">
				<%= FormatNumber(opgdataAdvPrice.FItemList(i).FreasonGubun003, 0) %>
			</a>
			<% end if %>
		</td>
		<td align=right>
			<% if Not IsNull(opgdataAdvPrice.FItemList(i).FreasonGubun020) then %>
			<a href="javascript:popMeachulPriceReasonDetail('<%= opgdataAdvPrice.FItemList(i).Fyyyymm %>', '<%= opgdataAdvPrice.FItemList(i).FtargetGbn %>', '<%= opgdataAdvPrice.FItemList(i).FPGgubun %>', '<%= opgdataAdvPrice.FItemList(i).FPGuserid %>', '020');">
				<%= FormatNumber(opgdataAdvPrice.FItemList(i).FreasonGubun020, 0) %>
			</a>
			<% end if %>
		</td>
		<td align=right>
			<% if Not IsNull(opgdataAdvPrice.FItemList(i).FreasonGubun025) then %>
			<a href="javascript:popMeachulPriceReasonDetail('<%= opgdataAdvPrice.FItemList(i).Fyyyymm %>', '<%= opgdataAdvPrice.FItemList(i).FtargetGbn %>', '<%= opgdataAdvPrice.FItemList(i).FPGgubun %>', '<%= opgdataAdvPrice.FItemList(i).FPGuserid %>', '025');">
				<%= FormatNumber(opgdataAdvPrice.FItemList(i).FreasonGubun025, 0) %>
			</a>
			<% end if %>
		</td>
		<td align=right>
			<% if Not IsNull(opgdataAdvPrice.FItemList(i).FreasonGubun030) then %>
			<a href="javascript:popMeachulPriceReasonDetail('<%= opgdataAdvPrice.FItemList(i).Fyyyymm %>', '<%= opgdataAdvPrice.FItemList(i).FtargetGbn %>', '<%= opgdataAdvPrice.FItemList(i).FPGgubun %>', '<%= opgdataAdvPrice.FItemList(i).FPGuserid %>', '030');">
				<%= FormatNumber(opgdataAdvPrice.FItemList(i).FreasonGubun030, 0) %>
			</a>
			<% end if %>
		</td>
		<td align=right>
			<% if Not IsNull(opgdataAdvPrice.FItemList(i).FreasonGubun035) then %>
			<a href="javascript:popMeachulPriceReasonDetail('<%= opgdataAdvPrice.FItemList(i).Fyyyymm %>', '<%= opgdataAdvPrice.FItemList(i).FtargetGbn %>', '<%= opgdataAdvPrice.FItemList(i).FPGgubun %>', '<%= opgdataAdvPrice.FItemList(i).FPGuserid %>', '035');">
				<%= FormatNumber(opgdataAdvPrice.FItemList(i).FreasonGubun035, 0) %>
			</a>
			<% end if %>
		</td>
		<td align=right>
			<% if Not IsNull(opgdataAdvPrice.FItemList(i).FreasonGubun004) then %>
			<a href="javascript:popMeachulPriceReasonDetail('<%= opgdataAdvPrice.FItemList(i).Fyyyymm %>', '<%= opgdataAdvPrice.FItemList(i).FtargetGbn %>', '<%= opgdataAdvPrice.FItemList(i).FPGgubun %>', '<%= opgdataAdvPrice.FItemList(i).FPGuserid %>', '004');">
				<%= FormatNumber(opgdataAdvPrice.FItemList(i).FreasonGubun004, 0) %>
			</a>
			<% end if %>
		</td>
		<td align=right>
			<% if Not IsNull(opgdataAdvPrice.FItemList(i).FreasonGubun040) then %>
			<a href="javascript:popMeachulPriceReasonDetail('<%= opgdataAdvPrice.FItemList(i).Fyyyymm %>', '<%= opgdataAdvPrice.FItemList(i).FtargetGbn %>', '<%= opgdataAdvPrice.FItemList(i).FPGgubun %>', '<%= opgdataAdvPrice.FItemList(i).FPGuserid %>', '040');">
				<%= FormatNumber(opgdataAdvPrice.FItemList(i).FreasonGubun040, 0) %>
			</a>
			<% end if %>
		</td>
		<td align=right>
			<% if Not IsNull(opgdataAdvPrice.FItemList(i).FreasonGubun800) then %>
			<a href="javascript:popMeachulPriceReasonDetail('<%= opgdataAdvPrice.FItemList(i).Fyyyymm %>', '<%= opgdataAdvPrice.FItemList(i).FtargetGbn %>', '<%= opgdataAdvPrice.FItemList(i).FPGgubun %>', '<%= opgdataAdvPrice.FItemList(i).FPGuserid %>', '800');">
				<%= FormatNumber(opgdataAdvPrice.FItemList(i).FreasonGubun800, 0) %>
			</a>
			<% end if %>
		</td>
		<td align=right>
			<% if Not IsNull(opgdataAdvPrice.FItemList(i).FreasonGubun900) then %>
			<a href="javascript:popMeachulPriceReasonDetail('<%= opgdataAdvPrice.FItemList(i).Fyyyymm %>', '<%= opgdataAdvPrice.FItemList(i).FtargetGbn %>', '<%= opgdataAdvPrice.FItemList(i).FPGgubun %>', '<%= opgdataAdvPrice.FItemList(i).FPGuserid %>', '900');">
				<%= FormatNumber(opgdataAdvPrice.FItemList(i).FreasonGubun900, 0) %>
			</a>
			<% end if %>
		</td>
		<td align=right>
			<% if Not IsNull(opgdataAdvPrice.FItemList(i).FreasonGubun901) then %>
			<a href="javascript:popMeachulPriceReasonDetail('<%= opgdataAdvPrice.FItemList(i).Fyyyymm %>', '<%= opgdataAdvPrice.FItemList(i).FtargetGbn %>', '<%= opgdataAdvPrice.FItemList(i).FPGgubun %>', '<%= opgdataAdvPrice.FItemList(i).FPGuserid %>', '901');">
				<%= FormatNumber(opgdataAdvPrice.FItemList(i).FreasonGubun901, 0) %>
			</a>
			<% end if %>
		</td>
		<td align=right>
			<% if Not IsNull(opgdataAdvPrice.FItemList(i).FreasonGubun950) then %>
			<a href="javascript:popMeachulPriceReasonDetail('<%= opgdataAdvPrice.FItemList(i).Fyyyymm %>', '<%= opgdataAdvPrice.FItemList(i).FtargetGbn %>', '<%= opgdataAdvPrice.FItemList(i).FPGgubun %>', '<%= opgdataAdvPrice.FItemList(i).FPGuserid %>', '950');">
				<%= FormatNumber(opgdataAdvPrice.FItemList(i).FreasonGubun950, 0) %>
			</a>
			<% end if %>
		</td>
		<td align=right>
			<% if Not IsNull(opgdataAdvPrice.FItemList(i).FreasonGubun999) then %>
			<a href="javascript:popMeachulPriceReasonDetail('<%= opgdataAdvPrice.FItemList(i).Fyyyymm %>', '<%= opgdataAdvPrice.FItemList(i).FtargetGbn %>', '<%= opgdataAdvPrice.FItemList(i).FPGgubun %>', '<%= opgdataAdvPrice.FItemList(i).FPGuserid %>', '999');">
				<%= FormatNumber(opgdataAdvPrice.FItemList(i).FreasonGubun999, 0) %>
			</a>
			<% end if %>
		</td>
		<td align=right>
			<% if Not IsNull(opgdataAdvPrice.FItemList(i).FreasonGubunXXX) then %>
			<a href="javascript:popMeachulPriceReasonDetail('<%= opgdataAdvPrice.FItemList(i).Fyyyymm %>', '<%= opgdataAdvPrice.FItemList(i).FtargetGbn %>', '<%= opgdataAdvPrice.FItemList(i).FPGgubun %>', '<%= opgdataAdvPrice.FItemList(i).FPGuserid %>', 'XXX');">
				<%= FormatNumber(opgdataAdvPrice.FItemList(i).FreasonGubunXXX, 0) %>
			</a>
			<% end if %>
		</td>
		<td align=center><%= opgdataAdvPrice.FItemList(i).Fregdate %></td>
		<td>
	    </td>
	</tr>
	<% next %>
<% else %>
	<tr bgcolor="#FFFFFF" height="25">
		<td colspan="26" align=center>[ �˻������ �����ϴ�. ]</td>
	</tr>
<% end if %>
</table>

<form id="frmAct" name="frmAct" method="post" action="https://scm.10x10.co.kr/admin/maechul/pgdata/pgdata_process.asp">
	<input type="hidden" name="mode" value="">
	<input type="hidden" name="yyyymm" value="<%= yyyy1 %>-<%= mm1 %>">
</form>

<%
set opgdataAdvPrice = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
