<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
' Description : �Ǹų���[Ư����ǰ] �����ٿ�ε�
' History	:  2022.09.19 �ѿ�� ����
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/items/new_itemcls.asp"-->
<!-- #include virtual="/lib/classes/maechul/managementSupport/maechulCls.asp" -->
<%
dim itemid, itemoption, itemstate, sitename, inccancel, yyyy1,yyyy2,mm1,mm2,dd1,dd2, nowdate,oldlist
dim a1010,w1010,m1010,w10102,a10102,m10102, premonthdate, datetype, sortType, RowArr, oItemOrder
dim fromDate,toDate, page, RowCount, jumuncnt, totno, i, oitem, oitemoption
dim itemnosum, sellprice, realsellprice, upchejungsanprice
    page = RequestCheckVar(getNumeric(trim(request("page"))),10)
	nowdate         = Left(CStr(now()),10)
	premonthdate    = DateAdd("d",-14,nowdate)
	itemid = requestCheckvar(getNumeric(trim(request("itemid"))),10)
	itemoption = requestCheckvar(request("itemoption"),10)
	itemstate = request("itemstate")
	oldlist = request("oldlist")
	yyyy1   = requestCheckvar(getNumeric(trim(request("yyyy1"))),4)
	mm1     = requestCheckvar(getNumeric(trim(request("mm1"))),2)
	dd1     = requestCheckvar(getNumeric(trim(request("dd1"))),2)
	yyyy2   = requestCheckvar(getNumeric(trim(request("yyyy2"))),4)
	mm2     = requestCheckvar(getNumeric(trim(request("mm2"))),2)
	dd2     = requestCheckvar(getNumeric(trim(request("dd2"))),2)
	datetype = request("datetype")
	sitename = requestCheckvar(request("sitename"),32)
	inccancel = requestCheckvar(request("inccancel"),1)
	a1010 = requestCheckvar(request("a1010"),10)
	w1010 = requestCheckvar(request("w1010"),1)
	m1010 = requestCheckvar(request("m1010"),10)
	sortType = requestCheckvar(request("sortType"),2)

if sortType="" then sortType="od"
if (itemstate="5") then itemstate="6"
if (yyyy1="") then
	yyyy1 = Left(premonthdate,4)
	mm1   = Mid(premonthdate,6,2)
	dd1   = Mid(premonthdate,9,2)

	nowdate = Left(CStr(now()),10)
	yyyy2 = Left(nowdate,4)
	mm2   = Mid(nowdate,6,2)
	dd2   = Mid(nowdate,9,2)
else
	nowdate = Left(CStr(DateSerial(yyyy1 , mm1 , dd1)),10)
	yyyy1 = Left(nowdate,4)
	mm1   = Mid(nowdate,6,2)
	dd1   = Mid(nowdate,9,2)
end if
if (page="") then page=1
fromDate = CStr(DateSerial(yyyy1, mm1, dd1))
toDate = CStr(DateSerial(yyyy2, mm2, dd2+1))

if (datetype="") then datetype="reg"

if w1010 <> "" or m1010 <> "" or a1010 <> "" then
	if w1010="Y" then
		w10102=""
	else
		w10102="N"
	end if
	if m1010="" then
		m10102="N"
	else
		m10102=m1010
	end if
	if a1010="" then
		a10102="N"
	else
		a10102=a1010
	end if
end if

'��ǰ�ڵ� ��ȿ�� �˻�(2008.08.05;������)
if itemid<>"" then
	if Not(isNumeric(itemid)) then
		Response.Write "<script type='text/javascript'>alert('[" & itemid & "]��(��) ��ȿ�� ��ǰ�ڵ尡 �ƴմϴ�.');history.back();</script>"
		dbget.close()	:	response.End
	end if
end if

set oItemOrder = new cManagementSupportMaechul_list
	oItemOrder.FCurrPage = page
	oItemOrder.FPageSize = 200000
	oItemOrder.FRectStartDate = fromDate
	oItemOrder.FRectEndDate   = toDate
	oItemOrder.frectdatetype=datetype
	oItemOrder.frectinccancel=inccancel
	oItemOrder.frectitemoption=itemoption
	oItemOrder.frectitemstate=itemstate
	oItemOrder.frectsitename=sitename
	oItemOrder.frectw10102=w10102
	oItemOrder.frectm10102=m10102
	oItemOrder.frecta10102=a10102

	if itemid<>"" and not(isnull(itemid)) then
		oItemOrder.GetOneItemOrderListNotPaging
	end if

if oItemOrder.FTotalCount>0 then
    RowArr=oItemOrder.fArrLIst
end if

RowCount = 0
jumuncnt = 0
if IsArray(RowArr) then
    RowCount = Ubound(RowArr,2)
    jumuncnt = RowCount + 1
end if

totno = 0

set oitem = new CItemInfo
oitem.FRectItemID = itemid

if itemid<>"" then
	oitem.GetOneItemInfo
end if

set oitemoption = new CItemOption
oitemoption.FRectItemID = itemid
if itemid<>"" then
	oitemoption.GetItemOptionInfo
end if

Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=TENITEMBUYLIST" & Left(CStr(now()),10) & "_" & session.sessionID & ".xls"
Response.CacheControl = "public"
Response.Buffer = true    '���ۻ�뿩��
%>
<html>
<head>
<style type='text/css'>
	.txt {mso-number-format:'\@'}
</style>
</head>
<body>
<table width="100%" align="center" cellpadding="3" cellspacing="1" border=1 bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
    <td>�ֹ���ȣ</td>
    <td>����</td>
    <td>���Ա���</td>
    <td>��������</td>
    <td>Site</td>
    <td>rdSite</td>
    <td>�ֹ�����</td>
    <td>��ǰ����</td>
    <td>����</td>
    <td>�ɼǸ�</td>
    <td>�ɼ��ڵ�</td>
    <td>ȸ��ID</td>
    <td>ȸ�����</td>
    <% if (FALSE) then %>
        <td>������</td>
    <% end if %>
    <td>������</td>
    <td>�ǸŰ�</td>
    <% if (C_InspectorUser = False) then %>
        <td>���ǸŰ�(��������)</td>
    <% end if %>
    <td>��ü�����</td>
    <td>�ֹ���</td>
    <td>�����</td>
    <td>�����</td>
    <td>������</td>
    <td>��������</td>
</tr>
<%
itemnosum = 0
sellprice = 0
realsellprice = 0
upchejungsanprice = 0

if IsArray(RowArr) then
    for i=0 to RowCount
    itemnosum = itemnosum + RowArr(2,i)
    sellprice = sellprice + RowArr(17,i)
    realsellprice = realsellprice + RowArr(18,i)
    upchejungsanprice = upchejungsanprice + RowArr(19,i)
%>
<tr bgcolor="#FFFFFF" align="center">
    <td><%= RowArr(0,i) %></td>
    <td><%= getJumundivName(RowArr(15,i)) %></td>
    <td><%= (RowArr(16,i)) %></td>
    <td><%= RowArr(24,i) %></td>
    <td><%= RowArr(12,i) %></td>
    <td><%= RowArr(22,i) %></td>
    <td><%= IpkumDivName(RowArr(1,i)) %></td>
    <td><%= getCurrstateName(RowArr(1,i),RowArr(11,i)) %></td>
    <td><%= RowArr(2,i) %></td>
    <td><%= DdotFormat(RowArr(10,i),20) %></td>
    <td class="txt"><%= RowArr(23,i) %></td>
    <td class="txt">
        <% if C_CriticInfoUserLV1 or C_CriticInfoUserLV2 or C_CriticInfoUserLV3 then %>
            <%= RowArr(14,i) %>
        <% else %>
            <%= printUserId(RowArr(14,i),2,"*") %>
        <% end if %>
    </td>
    <td><%= getUserLevelStrByDate(RowArr(25,i), left(RowArr(21,i),10)) %></td>
    <% if (FALSE) then %>
    <td><%= RowArr(3,i) %></td>
    <% end if %>
    <td><%= RowArr(7,i) %></td>
    <% if (C_InspectorUser = False) then %>
    <td><%= FormatNumber(RowArr(17,i),0) %></td>
    <% end if %>
    <td><%= FormatNumber(RowArr(18,i),0) %></td>
    <td><%= FormatNumber(RowArr(19,i),0) %></td>
    <td><%= RowArr(21,i) %></td>
    <td><%= RowArr(13,i) %></td>
    <td><%= RowArr(28,i) %></td>
    <td><%= RowArr(29,i) %></td>
    <td><%= GetaccountdivName(RowArr(26,i)) %></td>
</tr>
<%
    totno = totno + RowArr(2,i)

    if i mod 1000 = 0 then
        Response.Flush		' ���۸��÷���
    end if
next
%>
<tr align="center" bgcolor="#FFFFFF">
    <td colspan=8>�Ѿ�</td>
    <td><%= FormatNumber(itemnosum,0) %></td>
    <td colspan=5>&nbsp;</td>

    <% if (C_InspectorUser = False) then %>
        <td><%= FormatNumber(sellprice,0) %></td>
    <% end if %>

    <td><%= FormatNumber(realsellprice,0) %></td>
    <td><%= FormatNumber(upchejungsanprice,0) %></td>
    <td colspan=5>&nbsp;</td>
</tr>
<tr bgcolor="FFFFFF" align="center">
	<td colspan="22">
		�ѻ�ǰ�� <%= totno %> �� / ���ֹ��Ǽ� <%= jumuncnt %> ��
	</td>
</tr>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td colspan="22" align="center">[�˻������ �����ϴ�.]</td>
	</tr>
<% end if %>
</table>
</body>
</html>
<%
function IpkumDivName(byval v )
	if v="0" then
		IpkumDivName="�ֹ����"
	elseif v="1" then
		IpkumDivName="�ֹ�����"
	elseif v="2" then
		IpkumDivName="�ֹ�����"
	elseif v="3" then
		IpkumDivName="�ֹ�����"
	elseif v="4" then
		IpkumDivName="�����Ϸ�"
	elseif v="5" then
		IpkumDivName="�ֹ��뺸"
	elseif v="6" then
		IpkumDivName="��ǰ�غ�"
	elseif v="7" then
		IpkumDivName="�Ϻ����"
	elseif v="8" then
		IpkumDivName="���Ϸ�"
	elseif v="9" then
		IpkumDivName="���̳ʽ�"
	end if
end function

function getCurrstateName(byval v1, byval v)
    if (v=0) then
        if (v1>3) and (v1<8) then
           getCurrstateName = "�����Ϸ�"
        else
            getCurrstateName = IpkumDivName(v1)
        end if
    else
        if v=2 then
            getCurrstateName = "�ֹ��뺸"
        elseif v=3 then
            getCurrstateName = "��ǰ�غ�"
        elseif v=7 then
            getCurrstateName = "���Ϸ�"
        else
            getCurrstateName = v
        end if
    end if
end function

function getCurrstateNameColor(byval v1, byval v)
    if (v=0) then
        if (v1>3) and (v1<8) then
            getCurrstateNameColor = IpkumDivColor(4)
        else
            getCurrstateNameColor = IpkumDivName(v1)
        end if
    else
        if v=2 then
            getCurrstateNameColor = IpkumDivColor(v)
        elseif v=3 then
            getCurrstateNameColor = IpkumDivColor(v)
        elseif v=7 then
            getCurrstateNameColor = IpkumDivColor(v)
        else
            getCurrstateNameColor = "#000000"
        end if
    end if
end function

function getJumundivName(byval ijumundiv)
    if (isNULL(ijumundiv)) then
        getJumundivName = ""
        Exit function
    end if

    if ijumundiv="1" then
		getJumundivName="���"
	elseif ijumundiv="5" then
	    getJumundivName="���"
    elseif ijumundiv="9" then
        getJumundivName="<font color='red'>��ǰ</font>"
    elseif ijumundiv="6" then
        getJumundivName="<font color='blue'>��ȯ</font>"
    else
        getJumundivName=ijumundiv
    end if
end function

Function pointUpDown(txt,tp,sw,ud)
	dim ret, st
	st = tp & chkIIF(sw and ud,"a","d")
	ret = "<div class=""sorting"" style=""" & chkIIF(sw,"font-weight:bold;","") & """ onClick=""chgSortType('" & st & "')"">"
	ret = ret & txt
	ret = ret & "<span class=""" & chkIIF(sw and ud,"sortWay","") & """></span>"
	ret = ret & "</div>"
	pointUpDown = ret
end function

set oitem = Nothing
set oitemoption = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->