<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ���޸� Ŭ����
' Hieditor : 2011.04.22 �̻� ����
'			 2012.08.24 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionSTAdmin.asp" -->
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/maechul/pgdatacls.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%

dim research, page
dim excmatchfinish, onlyCardPriceNotSame, excChargeInput, pggubun
dim yyyy1,yyyy2,mm1,mm2,dd1,dd2
dim yyyy, mm, dd
dim fromDate ,toDate, tmpDate
dim shopid
dim appDivCode, cardReaderID, cardGubun, cardComp, cardAffiliateNo, ipkumdate
dim searchfield, searchtext, dateType
dim reasonGubun, PGuserid

Dim i

	research = requestCheckvar(request("research"),10)
	page = requestCheckvar(request("page"),10)
	excmatchfinish = requestCheckvar(request("excmatchfinish"),10)
	excChargeInput = requestCheckvar(request("excChargeInput"),10)
	onlyCardPriceNotSame = requestCheckvar(request("onlyCardPriceNotSame"),10)
	pggubun 		= request("pggubun")

	yyyy1   = request("yyyy1")
	mm1     = request("mm1")
	dd1     = request("dd1")
	yyyy2   = request("yyyy2")
	mm2     = request("mm2")
	dd2     = request("dd2")

	shopid 			= request("shopid")
	appDivCode 		= request("appDivCode")
	cardReaderID 	= request("cardReaderID")
	cardGubun 		= request("cardGubun")
	cardComp 		= request("cardComp")
	cardAffiliateNo = request("cardAffiliateNo")
	ipkumdate 		= request("ipkumdate")
	dateType 		= request("dateType")
	reasonGubun 	= requestCheckvar(request("reasonGubun"),32)

	searchfield 	= request("searchfield")
	searchtext 		= Replace(Replace(request("searchtext"), "'", ""), Chr(34), "")

	if request("PGuserid") <> "" then
		PGuserid = request("PGuserid")

		select case PGuserid
			case "BC":
				PGuserid = "��ī���"
			case "LOTTE":
				PGuserid = "�Ե�ī���"
			case "SAMSUNG":
				PGuserid = "�Ｚī���"
			case "SHINHAN":
				PGuserid = "����ī��"
			case "HANACARD":
				PGuserid = "�ϳ�ī��"
			case "HYUNDAI":
				PGuserid = "����ī���"
			case "ALI":
				PGuserid = "Alipay"
			case "KB":
				PGuserid = "KB����ī��"
			case "NH":
				PGuserid = "NH����ī��"
			case else:
				'//
		end select

		cardComp = PGuserid
	end if

if (page="") then page = 1

if (yyyy1="") then
	fromDate = DateSerial(Cstr(Year(now())), Cstr(Month(now()) - 1), 1)
	toDate = DateSerial(Cstr(Year(now())), Cstr(Month(now())) + 1, 1)

	yyyy1 = Cstr(Year(fromDate))
	mm1 = Cstr(Month(fromDate))
	dd1 = Cstr(day(fromDate))

	tmpDate = DateAdd("d", -1, toDate)
	yyyy2 = Cstr(Year(tmpDate))
	mm2 = Cstr(Month(tmpDate))
	dd2 = Cstr(day(tmpDate))
else
	fromDate = DateSerial(yyyy1, mm1, dd1)
	toDate = DateSerial(yyyy2, mm2, dd2+1)
end if


Dim oCPGData
set oCPGData = new CPGData
	oCPGData.FPageSize = 20
	oCPGData.FCurrPage = page

	oCPGData.FRectExcMatchFinish   	= excmatchfinish
	oCPGData.FRectExcChargeInput   	= excChargeInput


	oCPGData.FRectDateType = dateType
	oCPGData.FRectStartdate = fromDate
	oCPGData.FRectEndDate = toDate

	oCPGData.FRectshopid = shopid
	oCPGData.FRectAppDivCode = appDivCode
	oCPGData.FRectPGGubun = pggubun
	oCPGData.FRectCardReaderID = cardReaderID
	oCPGData.FRectCardGubun = cardGubun
	oCPGData.FRectCardComp = cardComp
	oCPGData.FRectCardAffiliateNo = cardAffiliateNo
	oCPGData.FRectIpkumdate = ipkumdate
	oCPGData.FRectReasonGubun = reasonGubun

	oCPGData.FRectSearchField = searchfield
	oCPGData.FRectSearchText = searchtext
	oCPGData.FRectOnlyCardPriceNotSame = onlyCardPriceNotSame

    oCPGData.getPGDataList_OFF

%>

<script language='javascript'>

function NextPage(page){
    document.frm.page.value = page;
    document.frm.submit();
}

function popUploadPGData() {
    var window_width = 500;
    var window_height = 250;

    var popwin = window.open("popRegPGDataFile_off.asp","popUploadPGData","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");

	popwin.focus();
}

function popUploadHandData() {
	var popwin = window.open("popRegHand_off.asp","popUploadHandData","width=600 height=200 scrollbars=yes resizable=yes status=yes");

	popwin.focus();
}

function popUploadPGChargeData() {
    var window_width = 500;
    var window_height = 250;

    var popwin = window.open("popRegPGChargeDataFile_off.asp","popUploadPGChargeData","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");

	popwin.focus();
}

function popjumundetail(yyyy1, mm1, dd1, shopid, logidx, cardsum) {
	var popjumundetail = window.open("popOffShopOrderList.asp?menupos=648&oldlist=&datefg=jumun&yyyy1="+yyyy1+"&mm1="+mm1+"&dd1="+dd1+"&yyyy2="+yyyy1+"&mm2="+mm1+"&dd2="+dd1+"&shopid="+shopid+"&buyergubun=" + "&logidx=" + logidx + "&cardsum=" + cardsum,"popjumundetail","width=1024,height=768,scrollbars=yes,resizable=yes");
	popjumundetail.focus();
}

function jsDelMatch(logidx) {
	var frm = document.frmAct;

	frm.logidx.value = logidx;

	if (confirm("��Ī[����] �Ͻðڽ��ϱ�?") == true) {
		frm.submit();
	}
}

function jsMatchCancel(logidx) {
	var frm = document.frmAct;

	frm.logidx.value = logidx;
	frm.mode.value = "matchcancel";

	if (confirm("[���]���� ��Ī �Ͻðڽ��ϱ�?") == true) {
		frm.submit();
	}
}

function jsPopRegReasonGubun(idx) {
	var v = "popRegReasonGubun.asp?idx=" + idx + "&gubun=off";
	var popwin = window.open(v,"jsPopRegReasonGubun","width=250,height=150,scrollbars=yes,resizable=yes");
	popwin.focus();
}

</script>
<link rel="stylesheet" href="/css/tpl.css" type="text/css">

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		&nbsp;
		<select class="select" name="dateType">
			<option value="A" <% if (dateType = "A") then %>selected<% end if %> >�ŷ���</option>
			<option value="B" <% if (dateType = "B") then %>selected<% end if %> >�Աݿ�����</option>
		</select>
		:
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		&nbsp;
		* ���α��� :
		<select class="select" name="appDivCode">
		<option value=""></option>
		<option value="A" <% if (appDivCode = "A") then %>selected<% end if %> >����</option>
		<option value="C" <% if (appDivCode = "C") then %>selected<% end if %> >���</option>
		<option value="P" <% if (appDivCode = "P") then %>selected<% end if %> >�������</option>
		</select>
		&nbsp;
		* �ܸ����ȣ :
		<input type="text" class="text" name="cardReaderID" value="<%= cardReaderID %>" size="8">
		&nbsp;
		* ī�屸�� :
		<select class="select" name="cardGubun">
		<option value=""></option>
		<option value="�ſ�" <% if (cardGubun = "�ſ�") then %>selected<% end if %> >�ſ�</option>
		<option value="üũ" <% if (cardGubun = "üũ") then %>selected<% end if %> >üũ</option>
		</select>
	</td>
	<td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td align="left">
		&nbsp;
		* PG�� :
		<select class="select" name="pggubun">
		<option value=""></option>
		<option value="KICC" <% if (pggubun = "KICC") then %>selected<% end if %> >KICC</option>
		<option value="HAND" <% if (pggubun = "HAND") then %>selected<% end if %> >����</option>
		</select>
		&nbsp;
		* ī��� :
		<input type="text" class="text" name="cardComp" value="<%= cardComp %>" size="10">
		&nbsp;
		* ��������ȣ :
		<input type="text" class="text" name="cardAffiliateNo" value="<%= cardAffiliateNo %>" size="10">
		&nbsp;
		* �Աݿ����� :
		<input type="text" class="text" name="ipkumdate" value="<%= ipkumdate %>" size="10">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td align="left">
		&nbsp;
		* ���� : <% drawSelectBoxOffShopdiv_off "shopid", shopid, "1,3,7,9,11", "", "" %>
		&nbsp;
		* �˻����� :
		<select class="select" name="searchfield">
		<option value=""></option>
		<option value="PGkey" <% if (searchfield = "PGkey") then %>selected<% end if %> >PG��KEY</option>
		<option value="cardPrice" <% if (searchfield = "cardPrice") then %>selected<% end if %> >�ŷ���</option>
		<option value="cardAppNo" <% if (searchfield = "cardAppNo") then %>selected<% end if %> >���ι�ȣ</option>
		<option value="orderserial" <% if (searchfield = "orderserial") then %>selected<% end if %> >�ֹ���ȣ</option>
		<option value="orderCardPrice" <% if (searchfield = "orderCardPrice") then %>selected<% end if %> >�ֹ�ī���</option>
		<option value="ipkumPrice" <% if (searchfield = "ipkumPrice") then %>selected<% end if %> >�Աݿ�����</option>
		</select>
		<input type="text" class="text" name="searchtext" value="<%= searchtext %>">
		&nbsp;
		* �󼼻��� :
		<select class="select" name="reasonGubun">
		<option value=""></option>
		<option value="001" <% if (reasonGubun = "001") then %>selected<% end if %> >������(����)</option>
		<option value="002" <% if (reasonGubun = "002") then %>selected<% end if %> >������(���޻� ����)</option>
		<option value="020" <% if (reasonGubun = "020") then %>selected<% end if %> >������(��ġ��)</option>
		<option value="025" <% if (reasonGubun = "025") then %>selected<% end if %> >������(��ġ��ȯ��)</option>
		<option value="030" <% if (reasonGubun = "030") then %>selected<% end if %> >������(����Ʈ)</option>
		<option value="035" <% if (reasonGubun = "035") then %>selected<% end if %> >������(����Ʈȯ��)</option>
		<option value="">---------------</option>
		<option value="040" <% if (reasonGubun = "040") then %>selected<% end if %> >CS����</option>
		<option value="">---------------</option>
		<option value="950" <% if (reasonGubun = "950") then %>selected<% end if %> >�������Ȯ��</option>
		<option value="999" <% if (reasonGubun = "999") then %>selected<% end if %> >��Ҹ�Ī</option>
		<option value="901" <% if (reasonGubun = "901") then %>selected<% end if %> >�ΰŽ����ݸ���</option>
		<option value="800" <% if (reasonGubun = "800") then %>selected<% end if %> >���ڼ���</option>
		<option value="900" <% if (reasonGubun = "900") then %>selected<% end if %> >��Ÿ</option>
		<option value="">---------------</option>
		<option value="XXX" <% if (reasonGubun = "XXX") then %>selected<% end if %> >�Է�����</option>
		</select>
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td align="left">
		&nbsp;
		<input type="checkbox" name="excmatchfinish"  value="Y" <% if (excmatchfinish = "Y") then %>checked<% end if %> > �ֹ����� ��Ī�Ϸ� ����
		&nbsp;
		<input type="checkbox" name="onlyCardPriceNotSame"  value="Y" <% if (onlyCardPriceNotSame = "Y") then %>checked<% end if %> > �����ݾ� ����ġ������
		&nbsp;
		<input type="checkbox" name="excChargeInput"  value="Y" <% if (excChargeInput = "Y") then %>checked<% end if %> > ������ �Է¿Ϸ� ����
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->

<p />

* �󼼻��� �Է½��� : PG�� �ڷ� �Է½�<br />
&nbsp;&nbsp; - PG�� �ڷ� �Է� �� �ֹ����� �Է��ϸ� �󼼻��� �Է¾ȵ�

<p />

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">
		<input type="button" class="button" value="����ϱ�(PG�� �ڷ�)" onClick="popUploadPGData();">
		&nbsp;
		<input type="button" class="button" value="����ϱ�(KICC ������ �ڷ�)" onClick="popUploadPGChargeData();">
		&nbsp;
		<input type="button" class="button" value="����ϱ�(����)" onClick="popUploadHandData();">
	</td>
	<td align="right">
	</td>
</tr>
</table>
<!-- �׼� �� -->

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		�˻���� : <b><%= oCPGData.FTotalcount %></b>
		&nbsp;
		������ : <b><%= page %> / <%= oCPGData.FTotalPage %></b>
	</td>
</tr>
<form name="frm1" method="post">
<input type="hidden" name="mode" value="">
<input type="hidden" name="xSiteId" value="">
<input type="hidden" name="idx" value="">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="50">PG��</td>
	<td width="80">PG��KEY</td>
	<td width="60">����</td>
	<td width="150">�ŷ�����</td>
	<td width="60">�ܸ���<br>��ȣ</td>
	<td width="60">ī��<br>����</td>
	<td width="100">ī���</td>
	<td width="90">��������ȣ</td>
	<td width="60">�ŷ���</td>
	<td width="40">������</td>
	<td width="60">�Ա�<br>������</td>
	<td width="70">���ι�ȣ</td>
	<td width="70">�Աݿ�����</td>
	<td width="80">����</td>
	<td width="100">�ֹ���ȣ</td>
	<td width="60">�ֹ�<br>ī�����</td>
	<td>�󼼻���</td>
	<!--
	<td width="80">�����</td>
	-->
	<td>���</td>
</tr>

<% for i=0 to oCPGData.FresultCount -1 %>
<%
yyyy = Left(oCPGData.FItemList(i).FappDate, 4)
mm = Right(Left(oCPGData.FItemList(i).FappDate, 7), 2)
dd = Right(Left(oCPGData.FItemList(i).FappDate, 10), 2)

%>
<tr align="center" bgcolor="FFFFFF" onmouseover=this.style.background="F1F1F1"; onmouseout=this.style.background="FFFFFF";>
	<td><%= oCPGData.FItemList(i).FPGgubun %></td>
	<td><%= oCPGData.FItemList(i).FPGkey %></td>
	<td>
		<font color="<%= oCPGData.FItemList(i).GetAppDivCodeColor %>"><%= oCPGData.FItemList(i).GetAppDivCodeName %></font>
	</td>
	<td><%= oCPGData.FItemList(i).FappDate %></td>
	<td><%= oCPGData.FItemList(i).FcardReaderID %></td>
	<td><%= oCPGData.FItemList(i).FcardGubun %></td>
	<td><%= oCPGData.FItemList(i).FcardComp %></td>
	<td><%= oCPGData.FItemList(i).FcardAffiliateNo %></td>
	<td align="right"><%= FormatNumber(oCPGData.FItemList(i).FcardPrice, 0) %></td>
	<td align="right">
		<% if Not IsNull(oCPGData.FItemList(i).FcardChargePrice) then %>
			<%= FormatNumber(oCPGData.FItemList(i).FcardChargePrice, 0) %>
		<% end if %>
	</td>
	<td align="right">
		<% if Not IsNull(oCPGData.FItemList(i).FipkumPrice) then %>
			<%= FormatNumber(oCPGData.FItemList(i).FipkumPrice, 0) %>
		<% end if %>
	</td>
	<td><%= oCPGData.FItemList(i).FcardAppNo %></td>
	<td><%= oCPGData.FItemList(i).Fipkumdate %></td>
	<td>
		<%= oCPGData.FItemList(i).Fshopid %>
		<% if IsNull(oCPGData.FItemList(i).Fshopid) then %>
			(<%= oCPGData.FItemList(i).FcardReaderID %>)
		<% end if %>
	</td>
	<td><%= oCPGData.FItemList(i).Forderserial %></td>
	<td align="right">
		<% if Not IsNull(oCPGData.FItemList(i).ForderCardPrice) then %>
			<% if (oCPGData.FItemList(i).FcardPrice <> oCPGData.FItemList(i).ForderCardPrice) then %><font color="red"><% end if %>
			<%= FormatNumber(oCPGData.FItemList(i).ForderCardPrice, 0) %>
		<% end if %>
	</td>
	<td><%= oCPGData.FItemList(i).GetReasonGubunName %></td>
	<!--
	<td><%= Left(oCPGData.FItemList(i).Fregdate, 10) %></td>
	-->
	<td>
		<% if IsNull(oCPGData.FItemList(i).Forderserial) then %>
			<input type="button" onclick="popjumundetail('<%= yyyy %>','<%= mm %>','<%= dd %>','<%= oCPGData.FItemList(i).Fshopid %>', <%= oCPGData.FItemList(i).Fidx %>, <%= oCPGData.FItemList(i).FcardPrice %>);" value="�˻�" class="button">
			<% if (oCPGData.FItemList(i).FappDivCode = "C") or (oCPGData.FItemList(i).FappDivCode = "P") then %>
				<input type="button" onclick="jsMatchCancel(<%= oCPGData.FItemList(i).Fidx %>);" value="��Ҹ�Ī" class="button">
			<% end if %>
		<% else %>
			<input type="button" onclick="jsDelMatch(<%= oCPGData.FItemList(i).Fidx %>);" value="��Ī����" class="button">
		<% end if %>
		<% if IsNull(oCPGData.FItemList(i).FreasonGubun) or Not (InStr("001,020,030,950", oCPGData.FItemList(i).FreasonGubun) > 0) or C_ADMIN_AUTH then %>
			<input type="button" class="button" value="����" onClick="jsPopRegReasonGubun(<%= oCPGData.FItemList(i).Fidx %>)">
			<% if (C_ADMIN_AUTH) then %>[������]
			<% end if %>
		<% end if %>
	</td>
</tr>
<% next %>

<tr height="25" bgcolor="FFFFFF">
	<td colspan="20" align="center">
		<% if oCPGData.HasPreScroll then %>
		<a href="javascript:NextPage('<%= oCPGData.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + oCPGData.StartScrollPage to oCPGData.FScrollCount + oCPGData.StartScrollPage - 1 %>
			<% if i>oCPGData.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oCPGData.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
</form>
</table>

<%
set oCPGData = Nothing
%>

<form name="frmAct" method="post" action="<%=stsAdmURL%>/admin/maechul/pgdata/pgdata_process.asp">
<input type="hidden" name="mode" value="delmatchone">
<input type="hidden" name="logidx" value="">
<input type="hidden" name="orderno" value="">
</form>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
