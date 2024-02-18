<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/new_itemcls.asp"-->
<!-- #include virtual="/lib/BarcodeFunction.asp"-->
<%

function CurrStateName(byval v)
	if v < "B006" then
		CurrStateName="����"
	elseif v = "B006" then
		CurrStateName="��üó���Ϸ�"
	elseif v = "B007" then
		CurrStateName="ó���Ϸ�"
	else
		CurrStateName = v
	end if
end function

function CurrStateColor(byval v)
	if v < "B006" then
		CurrStateColor="#000000"
	elseif v = "B006" then
		CurrStateColor="#000000"
	elseif v = "B007" then
		CurrStateColor="green"
	else
		CurrStateColor = "gray"
	end if
end function

function DivcdName(byval v)
	if v = "A004" or v = "A010" then
		DivcdName="��ǰ"
	elseif v = "A000" then
		DivcdName="�±�ȯ���"
    elseif v = "A100" then
        DivcdName="��ȯ���"
	elseif v = "A002" then
		DivcdName="����"
	elseif v = "A011" then
	    DivcdName="�±�ȯȸ��"
	elseif v = "A012" or v = "A111" or v = "A112" then
		DivcdName="��ȯȸ��"
	elseif v = "CHG" then
	    DivcdName="��ȯCS"
	else
		DivcdName = v
	end if
end function


Const MaxRowSize = 1000
dim itemid, itemoption, itemgubun
dim currstate

dim datetype
dim startdate, enddate

itemid = request("itemid")
itemoption = request("itemoption")
currstate = request("currstate")

startdate = requestcheckvar(request("startdate"),10)
enddate = requestcheckvar(request("enddate"),10)

if startdate="" then
	startdate = Left(CStr(DateSerial(year(now), month(now), 1)),10)
end if
if enddate="" then
	enddate = date()
end if

datetype = request("datetype")

if (datetype="") then datetype="reg"
if (itemgubun="") then itemgubun="10"

datetype = "finish"
currstate = "finish"
''itemid = ""
''itemoption = ""


'��ǰ�ڵ� ��ȿ�� �˻�(2008.08.05;������)
if itemid<>"" then
	if Not(isNumeric(itemid)) then
		Response.Write "<script language=javascript>alert('[" & itemid & "]��(��) ��ȿ�� ��ǰ�ڵ尡 �ƴմϴ�.');history.back();</script>"
		dbget.close()	:	response.End
	end if
end if

dim sqlStr, RowArr

'[�ڵ�����]
'------------------------------------------------------------------------------
'A008			�ֹ����
'
'A004			��ǰ����(��ü���)
'A010			ȸ����û(�ٹ����ٹ��)
'
'A001			������߼�
'A002			���񽺹߼�
'
'A200			��Ÿȸ��
'
'A000			�±�ȯ���
'A100			��ǰ���� �±�ȯ���
'
'A009			��Ÿ����
'A006			�������ǻ���
'A700			��ü��Ÿ����
'
'A003			ȯ��
'A005			�ܺθ�ȯ�ҿ�û
'A007			ī��,��ü,�޴�����ҿ�û
'
'A011			�±�ȯȸ��(�ٹ����ٹ��)
'A012			�±�ȯ��ǰ(��ü���)

'A111			��ǰ���� �±�ȯȸ��(�ٹ����ٹ��)
'A112			��ǰ���� �±�ȯ��ǰ(��ü���)

''���񽺹߼�, �±�ȯ��� : ���� CS���� ���̳ʽ�
''��ǰ���� ��ȯ��� : ���� CS���� ���̳ʽ�, ��ȯȸ�� �Ϸ�� ���� �÷���
'' - ���� �Ѱ� ȸ���Ǵ� ��쵵 ����ؾ� �Ѵ�.
''��Ÿȸ�� CS��� �������� �ʴ´�.
''6���� ���� �ֹ��� ��� �˻��� ���� �ʴ´�.

sqlStr = " select top " & CStr(MaxRowSize) & " T.orderserial, T.sitename, T.chulgodate, T.itemid, T.itemoption, T.itemname, T.itemoptionname "

if (itemid <> "") then
	sqlStr = sqlStr + " , T.itemcnt "
	sqlStr = sqlStr + " , T.avgipgoprice "
	sqlStr = sqlStr + " , T.itemcost "
else
	sqlStr = sqlStr + " , isNull(sum(T.itemcnt),0) as itemcnt "
	sqlStr = sqlStr + " , isNull(sum(T.avgipgoprice),0) as avgipgoprice "
	sqlStr = sqlStr + " , isNull(sum(T.itemcost),0) as itemcost "
end if

sqlStr = sqlStr + " 	, T.divcd "
sqlStr = sqlStr + " from "
sqlStr = sqlStr + " 	( "
sqlStr = sqlStr + " 		select "

if (itemid <> "") then
	sqlStr = sqlStr + " 			a.orderserial "
	sqlStr = sqlStr + " 			, '' as sitename "
	sqlStr = sqlStr + " 			, a.finishdate as chulgodate "
	sqlStr = sqlStr + " 			, (case when a.divcd in ('A000', 'A002', 'A100') then d.confirmitemno * -1 else d.confirmitemno end) as itemcnt "
	sqlStr = sqlStr + " 			, ((case when a.divcd in ('A000', 'A002', 'A100') then d.confirmitemno * -1 else d.confirmitemno end)*s.avgipgoprice) as avgipgoprice "
	sqlStr = sqlStr + " 			, ((case when a.divcd in ('A000', 'A002', 'A100') then d.confirmitemno * -1 else d.confirmitemno end)*d.itemcost) as itemcost "
else
	sqlStr = sqlStr + " 			'' as orderserial "
	sqlStr = sqlStr + " 			, '' as sitename "
	sqlStr = sqlStr + " 			, '' as chulgodate "
	sqlStr = sqlStr + " 			, isNull(sum(case when a.divcd in ('A000', 'A002', 'A100') then d.confirmitemno * -1 else d.confirmitemno end),0) as itemcnt "
	sqlStr = sqlStr + " 			, IsNull(sum((case when a.divcd in ('A000', 'A002', 'A100') then d.confirmitemno * -1 else d.confirmitemno end)*s.avgipgoprice),0) as avgipgoprice "
	sqlStr = sqlStr + " 			, IsNull(sum((case when a.divcd in ('A000', 'A002', 'A100') then d.confirmitemno * -1 else d.confirmitemno end)*d.itemcost),0) as itemcost "
end if

sqlStr = sqlStr + " 			, d.itemid "
sqlStr = sqlStr + " 			, d.itemoption "
sqlStr = sqlStr + " 			, d.itemname "
sqlStr = sqlStr + " 			, d.itemoptionname "
sqlStr = sqlStr + " 			, a.divcd "
sqlStr = sqlStr + " 		from "
sqlStr = sqlStr + " 			db_cs.dbo.tbl_new_as_list a "
sqlStr = sqlStr + " 			join db_cs.dbo.tbl_new_as_detail d "
sqlStr = sqlStr + " 			on "
sqlStr = sqlStr + " 				a.id = d.masterid "
''sqlStr = sqlStr + " 			join [db_order].[dbo].tbl_order_master m "
''sqlStr = sqlStr + " 			on "
''sqlStr = sqlStr + " 				m.orderserial = a.orderserial "
sqlStr = sqlStr + "				join [db_summary].[dbo].[tbl_monthly_accumulated_logisstock_summary] s "
sqlStr = sqlStr + "				on "
sqlStr = sqlStr + "					1 = 1 "
sqlStr = sqlStr + "					and s.yyyymm = '" & Left(startdate, 7) & "' "
sqlStr = sqlStr + "					and s.itemgubun = '10' "
sqlStr = sqlStr + "					and s.itemid = d.itemid "
sqlStr = sqlStr + "					and s.itemoption = d.itemoption "
sqlStr = sqlStr + "					and s.lastmwdiv = 'M' "
sqlStr = sqlStr + " 		where "
sqlStr = sqlStr + " 			1 = 1 "
sqlStr = sqlStr + " 			and a.deleteyn <> 'Y' "
sqlStr = sqlStr + " 			and a.id >= 2500000 "
sqlStr = sqlStr + " 			and a.requireupche <> 'Y' "
sqlStr = sqlStr + " 			and a.divcd not in ('A008', 'A006', 'A001', 'A900', 'A010', 'A002', 'A111', 'A200', 'A999') "
sqlStr = sqlStr + " 			and a.currstate = 'B007' "
sqlStr = sqlStr + " 			and a.finishdate >= '" & startdate & "' "
sqlStr = sqlStr + " 			and a.finishdate < '" & enddate & "' "

if (itemid <> "") then
	sqlStr = sqlStr + " 	and d.itemid = " & itemid & " "
	if (itemoption <> "") then
		sqlStr = sqlStr + " 	and d.itemoption = '" & itemoption & "' "
	end if
else
	sqlStr = sqlStr + " 		group by "
sqlStr = sqlStr + " 			d.itemid, d.itemoption, d.itemname, d.itemoptionname,  a.divcd "
	sqlStr = sqlStr + " 		having sum(case when a.divcd in ('A000', 'A002', 'A100') then d.confirmitemno * -1 else d.confirmitemno end) <> 0 "
end if

sqlStr = sqlStr + " 		union all "
sqlStr = sqlStr + " 		select "

if (itemid <> "") then
	sqlStr = sqlStr + " 			a.orderserial "
	sqlStr = sqlStr + " 			, '' as sitename "
	sqlStr = sqlStr + " 			, a1.finishdate as chulgodate "
	sqlStr = sqlStr + " 			, d.confirmitemno as itemcnt "
	sqlStr = sqlStr + " 			, d.confirmitemno*s.avgipgoprice as avgipgoprice "
	sqlStr = sqlStr + " 			, d.confirmitemno*d.itemcost as itemcost "
else
	sqlStr = sqlStr + " 			'' as orderserial "
	sqlStr = sqlStr + " 			, '' as sitename "
	sqlStr = sqlStr + " 			, '' as chulgodate "
	sqlStr = sqlStr + " 			, isNull(sum(d.confirmitemno),0) as itemcnt "
	sqlStr = sqlStr + " 			, IsNull(sum(d.confirmitemno*s.avgipgoprice),0) as avgipgoprice "
	sqlStr = sqlStr + " 			, IsNull(sum(d.confirmitemno*d.itemcost),0) as itemcost "
end if

sqlStr = sqlStr + " 			, d.itemid "
sqlStr = sqlStr + " 			, d.itemoption "
sqlStr = sqlStr + " 			, d.itemname "
sqlStr = sqlStr + " 			, d.itemoptionname "
sqlStr = sqlStr + " 			, a.divcd "
sqlStr = sqlStr + " 		from "
sqlStr = sqlStr + " 			db_cs.dbo.tbl_new_as_list a "
sqlStr = sqlStr + " 			join db_cs.dbo.tbl_new_as_list a1 "
sqlStr = sqlStr + " 			on "
sqlStr = sqlStr + " 				1 = 1 "
sqlStr = sqlStr + " 				and a.id = a1.refasid "
sqlStr = sqlStr + " 			join db_cs.dbo.tbl_new_as_detail d "
sqlStr = sqlStr + " 			on "
sqlStr = sqlStr + " 				a.id = d.masterid "
''sqlStr = sqlStr + " 			join [db_order].[dbo].tbl_order_master m "
''sqlStr = sqlStr + " 			on "
''sqlStr = sqlStr + " 				m.orderserial = a.orderserial "
sqlStr = sqlStr + "				join [db_summary].[dbo].[tbl_monthly_accumulated_logisstock_summary] s "
sqlStr = sqlStr + "				on "
sqlStr = sqlStr + "					1 = 1 "
sqlStr = sqlStr + "					and s.yyyymm = '" & Left(startdate, 7) & "' "
sqlStr = sqlStr + "					and s.itemgubun = '10' "
sqlStr = sqlStr + "					and s.itemid = d.itemid "
sqlStr = sqlStr + "					and s.itemoption = d.itemoption "
sqlStr = sqlStr + "					and s.lastmwdiv = 'M' "
sqlStr = sqlStr + " 		where "
sqlStr = sqlStr + " 			1 = 1 "
sqlStr = sqlStr + " 			and a.requireupche <> 'Y' "
sqlStr = sqlStr + " 			and a.divcd = 'A100' "
sqlStr = sqlStr + " 			and a1.deleteyn <> 'Y' "
sqlStr = sqlStr + " 			and a1.id >= 2500000 "
sqlStr = sqlStr + " 			and a1.currstate = 'B007' "
sqlStr = sqlStr + " 			and a1.finishdate >= '" & startdate & "' "
sqlStr = sqlStr + " 			and a1.finishdate < '" & enddate & "' "

if (itemid <> "") then
	sqlStr = sqlStr + " 	and d.itemid = " & itemid & " "
	if (itemoption <> "") then
		sqlStr = sqlStr + " 	and d.itemoption = '" & itemoption & "' "
	end if
else
	sqlStr = sqlStr + " 		group by "
sqlStr = sqlStr + " 			d.itemid, d.itemoption, d.itemname, d.itemoptionname, a.divcd "
	sqlStr = sqlStr + " 		having sum(d.confirmitemno) <> 0 "
end if

sqlStr = sqlStr + " 	) T "

if (itemid = "") then
	sqlStr = sqlStr + " group by "
	sqlStr = sqlStr + " 	T.orderserial, T.sitename, T.chulgodate, T.itemid, T.itemoption, T.itemname, T.itemoptionname, T.divcd "
	sqlStr = sqlStr + " having sum(T.itemcnt) <> 0 "
end if

sqlStr = sqlStr + " order by "
sqlStr = sqlStr + " 	T.itemid, T.itemoption, T.chulgodate desc "

''response.write sqlStr
''response.end



IF application("Svr_Info")="Dev" THEN
    rsget.Open sqlStr,dbget,1
    if not rsget.Eof then
        RowArr = rsget.getRows
    end if
    rsget.Close
ELSE
    db3_rsget.Open sqlStr,db3_dbget,1
    if not db3_rsget.Eof then
        RowArr = db3_rsget.getRows
    end if
    db3_rsget.Close
end if

dim RowCount, jumuncnt
RowCount = 0
jumuncnt = 0
if IsArray(RowArr) then
    RowCount = Ubound(RowArr,2)
    jumuncnt = RowCount + 1
end if

dim totno, i
totno = 0

%>
<script type="text/javascript" src="/js/jsCal/js/jscal2.js"></script>
<script type="text/javascript" src="/js/jsCal/js/lang/ko.js"></script>
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/jscal2.css" />
<link rel="stylesheet" type="text/css" href="/js/jsCal/css/border-radius.css" />

<script type="text/javascript">
function jsSubmitONE(itemid, itemoption) {
	var frm = document.frm;

	frm.itemid.value = itemid;
	frm.itemoption.value = itemoption;

	frm.submit();
}
</script>

<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr height="25" bgcolor="#FFFFFF">
	    <td align="center" width="50" rowspan="2" bgcolor="#EEEEEE">�˻�<br />����</td>
        <td>
			�˻��Ⱓ :
			<select class="select" name="datetype">
			    <option value="finish" <%= chkIIF(datetype="finish","selected","") %> >ó����</option>
			</select>
			<input id="sDt" name="startdate" value="<%=startdate%>" class="text" size="10" maxlength="10" />
			<img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="sDt_trigger" border="0" style="cursor:pointer" align="absmiddle" /> ~
			<input id="eDt" name="enddate" value="<%=enddate%>" class="text" size="10" maxlength="10" />
			<img src="http://webadmin.10x10.co.kr/images/calicon.gif" id="eDt_trigger" border="0" style="cursor:pointer" align="absmiddle" />
			<script>
				var CAL_Start = new Calendar({
					inputField : "sDt", trigger    : "sDt_trigger",
					onSelect: function() {
						var date = Calendar.intToDate(this.selection.get());
						CAL_End.args.min = date;
						CAL_End.redraw();
						this.hide();
					}, bottomBar: true, dateFormat: "%Y-%m-%d"
				});
				var CAL_End = new Calendar({
					inputField : "eDt", trigger    : "eDt_trigger",
					onSelect: function() {
						var date = Calendar.intToDate(this.selection.get());
						CAL_Start.args.max = date;
						CAL_Start.redraw();
						this.hide();
					}, bottomBar: true, dateFormat: "%Y-%m-%d"
				});
			</script>

			&nbsp;
			��ǰ�ڵ� : <input type="text" class="text" name="itemid" value="<%= itemid %>" size="12">
			&nbsp;
			�ɼ� <input type="text" class="text" name="itemoption" value="<%= itemoption %>" size="4">
        </td>
        <td align="center" width="50" rowspan="2" bgcolor="#EEEEEE">
			<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
		</td>
	</tr>
	<tr height="25" bgcolor="#FFFFFF">
        <td>
			CS���� : ó���Ϸ� (�ִ� <%= MaxRowSize %>�� ������ �˻��˴ϴ�.)
        </td>
	</tr>
	</form>
</table>
<!-- �˻� �� -->
<p>
�� �����ؿ� : �˻����ۿ�<br />
 - ���񽺹߼�, �±�ȯ��� : ���� CS���� ���̳ʽ�<br />
 - ��ǰ���� ��ȯ��� : ���� CS���� ���̳ʽ�, ��ȯȸ�� �Ϸ�� ���� �÷���<br />
 - ��ǰ���� ��ȯȸ�� : �Ϸ�� ��ȯ�ֹ��� �����ǹǷ� CS���� ��������.<br />
&nbsp;&nbsp; - ���� �Ѱ� ȸ���Ǵ� ��쵵 ���<br />
 - ��Ÿȸ�� : CS��� �������� �ʴ´�.<br />
 - ��ȯCS : ���ϻ�ǰ ��ȯ��� �� ȸ��, ��ǰ���� ��ȯ��� �� ȸ��<br />
<!--
 - 6���� ���� �ֹ��� ��� �˻��� ���� �ʴ´�.
-->

<p>

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="80">�ֹ���ȣ</td>
		<td width="100">CS����</td>
		<td width="30">����</td>
		<td width="80">��ǰ�ڵ�</td>
		<td width="50">�ɼ�</td>
		<td width="100">���ڵ�</td>
		<td width="300">��ǰ��</td>
		<td width="200">�ɼǸ�</td>
		<td width="80">�ǸŰ�</td>
		<td width="80">��ո��԰�</td>
		<td width="40">����</td>
		<td width="150">�����</td>
		<td>���</td>
	</tr>
<%
if IsArray(RowArr) then
	for i=0 to RowCount
%>

	<tr align="center" bgcolor="#FFFFFF" height="25">
		<td><%= RowArr(0,i) %></td>
		<td><%= DivcdName(RowArr(10,i)) %></td>
		<td>10</td>
		<td><a href="javascript:jsSubmitONE('<%=RowArr(3,i)%>', '<%=RowArr(4,i)%>');"><%=RowArr(3,i)%></a></td>
		<td><a href="javascript:jsSubmitONE('<%=RowArr(3,i)%>', '<%=RowArr(4,i)%>');"><%=RowArr(4,i)%></a></td>
		<td align="left"><%= BF_MakeTenBarcode("10", RowArr(3,i), RowArr(4,i)) %></td>
		<td align="left"><%= DdotFormat(RowArr(5,i),25) %></td>
		<td align="left"><%= DdotFormat(RowArr(6,i),15) %></td>
		<td><%= RowArr(9,i) %></td>
		<td><%= RowArr(8,i) %></td>
		<td><%= RowArr(7,i) %></td>
		<td><%= RowArr(2,i) %></td>
		<td></td>
	</tr>
<%
			totno = totno + RowArr(7,i)
    next
end if

%>
    <tr height="25" bgcolor="#FFFFFF">
        <td align="right" colspan="13">�ѻ�ǰ�� <%= totno %> �� / ���ֹ��Ǽ� <%= jumuncnt %> ��</td>
    </tr>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
