<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �¶��� �������� ���ϸ��� & ��ġ�� ���հ���
' History : 2013.11.12 �ѿ�� ����
'           2018.03.12 ������ - ���ϸ��� ���� �߰�(����/���θ��)
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrShop.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/mileage/combine_point_deposit_cls.asp" -->
<%
Dim i, yyyy1,mm1,yyyy2,mm2, fromDate ,toDate ,ocombine, srcGbn, targetGbn
	yyyy1   = requestcheckvar(request("yyyy1"),10)
	mm1     = requestcheckvar(request("mm1"),10)
	yyyy2   = requestcheckvar(request("yyyy2"),10)
	mm2     = requestcheckvar(request("mm2"),10)
	srcGbn     = requestcheckvar(request("srcGbn"),1)
	targetGbn     = requestcheckvar(request("targetGbn"),4)

if (yyyy1="") then yyyy1 = Cstr(Year( dateadd("m",-3,date()) ))
if (mm1="") then mm1 = Cstr(Month( dateadd("m",-3,date()) ))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (srcGbn="") then srcGbn="M"

fromDate = left(DateSerial(yyyy1, mm1,"01"),7)
toDate = left(DateSerial(yyyy2, mm2+1,"01"),7)

Set ocombine = New ccombine_point_deposit
	ocombine.FRectStartdate = fromDate
	ocombine.FRectEndDate = toDate
	ocombine.FRectsrcGbn = srcGbn
	ocombine.FRecttargetGbn = targetGbn
	ocombine.FPageSize = 500
	ocombine.FCurrPage	= 1
	ocombine.fcombine_point_deposit_month()

'�� ǥ�� ���з�
dim rowSpanNo: rowSpanNo=1
dim colSpanNo: colSpanNo=1
if srcGbn="M" then
	rowSpanNo=2
	colSpanNo=2
end if

%>

<script language="javascript">
function searchSubmit(){
	frm.submit();
}

function pop_detail_list(yyyy1, mm1, dd1, yyyy2, mm2, dd2, GbnCd){
	var pop_detail_list = window.open('/admin/maechul/managementsupport/combine_point_deposit_list.asp?yyyy1='+yyyy1+'&mm1='+mm1+'&dd1='+dd1+'&yyyy2='+yyyy2+'&mm2='+mm2+'&dd2='+dd2+'&srcGbn=<%=srcGbn%>&targetGbn=<%=targetGbn%>&GbnCd='+GbnCd+'&menupos=<%=menupos%>','pop_detail_list','width=1024,height=768,scrollbars=yes,resizable=yes');
	pop_detail_list.focus();
}

function refreshSmr(yyyymm){
    if (confirm(yyyymm+' ���ۼ� �Ͻðڽ��ϱ�?')){
        document.frmAct.mode.value="refreshpointDepositSummary";
        document.frmAct.yyyymm.value=yyyymm;
        document.frmAct.submit();
    }
}
</script>

<!-- �˻� ���� -->
<form name="frm" method="get" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="70" bgcolor="<%= adminColor("gray") %>">�˻�</td>
	<td align="left">
		<table class="a">
		<tr>
			<td height="25">
				* ��¥ : <% DrawYMBoxdynamic "yyyy1",yyyy1,"mm1",mm1,"" %> ~ <% DrawYMBoxdynamic "yyyy2",yyyy2,"mm2",mm2,"" %>
				<p>
				* ���� : <% drawoffshop_commoncode "srcGbn", srcGbn, "srcGbn", "MAIN", "", "  " %>
				&nbsp;&nbsp;
				* ä�� : <% drawoffshop_commoncode "targetGbn", targetGbn, "targetGbn", "MAIN", "", "  " %>
			</td>
		</tr>
	    </table>
	</td>
	<td width="110" bgcolor="<%= adminColor("gray") %>"><input type="button" class="button_s" value="�˻�" onClick="javascript:searchSubmit();"></td>
</tr>
</table>
</form>
<!-- �˻� �� -->
<Br>
<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
<tr>
	<td align="left">
	</td>
	<td align="right">
	</td>
</tr>
</table>
<!-- �׼� �� -->

<p />

* ��� ���ξ� = �ֹ��û�� + ��ǰ/������� + ��Ÿ<br />
* ��ġ��/���ϸ����� �α����̺��� ������ ������.<br />
* ��ġ��<br />
&nbsp; - ������ȯ�� : �̼��� ������ȯ�� ������(��ġ��ȯ��) �� �ݾ��� ��ġ�ؾ� �մϴ�.<br />
&nbsp; - �ֹ��û�� : �����α� ������� ��ġ�ݰ� �ݾ��� ��ġ�ؾ� �մϴ�.<br />

<p />

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		�˻���� : <b><%= ocombine.FresultCount %></b> �� �� 500�Ǳ��� �˻� �˴ϴ�.
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
    <td rowspan="<%=rowSpanNo%>">��¥</td>
    <td rowspan="<%=rowSpanNo%>">�����հ�(����)</td>
    <td rowspan="<%=rowSpanNo%>">�����հ�</td>
    <td rowspan="<%=rowSpanNo%>">�ֹ�����</td>
    <% if (srcGbn="M") then %>
    <td colspan="<%=colSpanNo%>">�̺�Ʈ����</td>
    <% elseif (srcGbn="G") then %>
    <td>�̺�Ʈ</td>
    <% else %>
    <td>��ǰ/�������</td>
    <% end if %>
    <% if (srcGbn="G") then %>
    <td rowspan="<%=rowSpanNo%>">����(��)����ȯ</td>
    <td rowspan="<%=rowSpanNo%>">����Ʈī����(+)</td>
    <% else %>
    <td rowspan="<%=rowSpanNo%>">��ǰ����</td>
    <td rowspan="<%=rowSpanNo%>">CS����</td>
    <% end if %>
    <% if (srcGbn="M") then %>
    <td rowspan="<%=rowSpanNo%>">����>�� ��ȯ</td>
    <% elseif (srcGbn="G") then %>
    <td rowspan="<%=rowSpanNo%>">����Ʈī����(-)</td>
    <% else %>
    <td rowspan="<%=rowSpanNo%>">����(��)����ȯ</td>
    <% end if %>
    <!-- ���-->
    <td rowspan="<%=rowSpanNo%>">�ֹ��û��</td>
    <% if (srcGbn="M") then %>
    <td rowspan="<%=rowSpanNo%>">��Ÿ���</td>
    <% elseif (srcGbn="G") then %>
    <td rowspan="<%=rowSpanNo%>">������ ȯ��</td>
    <% else %>
    <td rowspan="<%=rowSpanNo%>">������ ȯ��</td>
    <% end if %>
    <td rowspan="<%=rowSpanNo%>">ȸ��Ż��</td>
    <td rowspan="<%=rowSpanNo%>">�Ҹ�</td>
    <td rowspan="<%=rowSpanNo%>">��Ÿ</td>
	<td rowspan="<%=rowSpanNo%>"><b>������ξ�</b></td>
    <% if (C_ADMIN_AUTH) then %><td rowspan="<%=rowSpanNo%>">ACT</td><% end if %>
</tr>
<% if (srcGbn="M") then %>
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
	<td>����</td>
	<td>�̺�Ʈ</td>
</tr>
<% end if %>
<% if C_MngPowerUser or C_ADMIN_AUTH then %>
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
    <td>YYYYMM</td>
    <td>accpointsum</td>
	<td>pointsum</td>
	<td>ORD</td>
	<td colspan="<%=colSpanNo%>">GNE</td>
	<td>GNI</td>
	<td>GNC</td>
	<td>SFT</td>
	<td>SPO</td>
	<td>SPE</td>
	<td>RTD</td>
    <td>XPR</td>
	<td>ETC</td>
	<td></td>
	<td>
	    <img src="/images/icon_reload.gif" onClick="refreshSmr('<%= Left(dateAdd("m",-1,now()),7) %>');" style="cursor:pointer;" alt="���ۼ�" title="<%= Left(dateAdd("m",-1,now()),7) %> ���� ���ۼ�">
	    <% if (day(now())>5) then '// 2017-09-07, 20�Ͽ��� 5�Ϸ�, skyer9 %>
	    ,
	    <img src="/images/icon_reload.gif" onClick="refreshSmr('<%= Left(dateAdd("m",0,now()),7) %>');" style="cursor:pointer;" alt="���ۼ�" title="<%= Left(dateAdd("m",0,now()),7) %> ���� ���ۼ�">
	    <% end if %>
	</td>
</tr>
<% end if %>
<%
dim totETC, totGNE, totSPE, totGNC, totGNI, totSPO, totSFT, totRTD, totORD, totpointsum, totXPR, totGOE, totGPE
	totETC=0
	totGNE=0
	totGOE=0		'// �̺�Ʈ ���ϸ��� ��������(jucyo:1100)
	totGPE=0		'// �̺�Ʈ ���ϸ��� ���θ������(jucyo:1000)
	totSPE=0
	totGNC=0
	totGNI=0
	totSPO=0
	totSFT=0
	totRTD=0
	totORD=0
	totXPR=0
	totpointsum=0

Dim oPoint : oPoint=0
Dim PPoint : PPoint=0

if srcGbn="M" and toDate="2013-12" and targetGbn="ONAC" then
    oPoint =1092762032
end if

if srcGbn="M" and toDate="2013-12" and targetGbn="OF" then
    oPoint =99979777
end if

if srcGbn="M" and toDate="2013-12" and targetGbn="" then
    oPoint =99979777+1092762032
end if

if srcGbn="M" and toDate="2013-12" and targetGbn="AC" then
    oPoint =8887135+4590279
end if


if srcGbn="M" and toDate="2013-12" and targetGbn="ON" then
    oPoint =1092762032-(8887135+4590279)
end if


if ocombine.FresultCount > 0 then

For i = 0 To ocombine.FresultCount -1
totpointsum = totpointsum + ocombine.fitemlist(i).fpointsum
totETC = totETC + ocombine.fitemlist(i).fETC
totGNE = totGNE + ocombine.fitemlist(i).fGNE
totGOE = totGOE + ocombine.fitemlist(i).fGOE
totGPE = totGPE + ocombine.fitemlist(i).fGPE
totSPE = totSPE + ocombine.fitemlist(i).fSPE
totGNC = totGNC + ocombine.fitemlist(i).fGNC
totGNI = totGNI + ocombine.fitemlist(i).fGNI
totSPO = totSPO + ocombine.fitemlist(i).fSPO
totSFT = totSFT + ocombine.fitemlist(i).fSFT
totRTD = totRTD + ocombine.fitemlist(i).fRTD
totORD = totORD + ocombine.fitemlist(i).fORD
totXPR = totXPR + ocombine.fitemlist(i).fXPR

oPoint = oPoint-pPoint
%>
<tr bgcolor="#FFFFFF" align="center" onmouseover=this.style.background="#F1F1F1"; onmouseout=this.style.background='#FFFFFF';>
	<td height="25">
		<%= ocombine.fitemlist(i).fYYYYMM %>
	</td>
	<td align="right" bgcolor="#9DCFFF">
         <%= FormatNumber(ocombine.fitemlist(i).faccpointsum,0) %>
	</td>
	<td align="right" bgcolor="#E6B9B8">
        <%= FormatNumber(ocombine.fitemlist(i).fpointsum,0) %>
    </td>
	<td align="right">
		<a href="javascript:pop_detail_list('<%= left(ocombine.fitemlist(i).fYYYYMM,4) %>','<%= mid(ocombine.fitemlist(i).fYYYYMM,6,2) %>','01','<%= left(ocombine.fitemlist(i).fYYYYMM,4) %>','<%= mid(ocombine.fitemlist(i).fYYYYMM,6,2) %>','<%= LastDayOfThisMonth( left(ocombine.fitemlist(i).fYYYYMM,4),mid(ocombine.fitemlist(i).fYYYYMM,6,2)) %>','ORD');" onfocus="this.blur()">
		<%= FormatNumber(ocombine.fitemlist(i).fORD,0) %>
		</a>
	</td>
	<% if (srcGbn="M") then %>
	<td align="right">
		<a href="javascript:pop_detail_list('<%= left(ocombine.fitemlist(i).fYYYYMM,4) %>','<%= mid(ocombine.fitemlist(i).fYYYYMM,6,2) %>','01','<%= left(ocombine.fitemlist(i).fYYYYMM,4) %>','<%= mid(ocombine.fitemlist(i).fYYYYMM,6,2) %>','<%= LastDayOfThisMonth( left(ocombine.fitemlist(i).fYYYYMM,4),mid(ocombine.fitemlist(i).fYYYYMM,6,2)) %>','GOE');" onfocus="this.blur()">
		<%= FormatNumber(ocombine.fitemlist(i).fGOE,0) %>
		</a>
	</td>
	<td align="right">
		<a href="javascript:pop_detail_list('<%= left(ocombine.fitemlist(i).fYYYYMM,4) %>','<%= mid(ocombine.fitemlist(i).fYYYYMM,6,2) %>','01','<%= left(ocombine.fitemlist(i).fYYYYMM,4) %>','<%= mid(ocombine.fitemlist(i).fYYYYMM,6,2) %>','<%= LastDayOfThisMonth( left(ocombine.fitemlist(i).fYYYYMM,4),mid(ocombine.fitemlist(i).fYYYYMM,6,2)) %>','GPE');" onfocus="this.blur()">
		<%= FormatNumber(ocombine.fitemlist(i).fGPE,0) %>
		</a>
	</td>
	<% else %>
	<td align="right">
		<a href="javascript:pop_detail_list('<%= left(ocombine.fitemlist(i).fYYYYMM,4) %>','<%= mid(ocombine.fitemlist(i).fYYYYMM,6,2) %>','01','<%= left(ocombine.fitemlist(i).fYYYYMM,4) %>','<%= mid(ocombine.fitemlist(i).fYYYYMM,6,2) %>','<%= LastDayOfThisMonth( left(ocombine.fitemlist(i).fYYYYMM,4),mid(ocombine.fitemlist(i).fYYYYMM,6,2)) %>','GNE');" onfocus="this.blur()">
		<%= FormatNumber(ocombine.fitemlist(i).fGNE,0) %>
		</a>
	</td>
	<% end if %>
	<td align="right">
		<a href="javascript:pop_detail_list('<%= left(ocombine.fitemlist(i).fYYYYMM,4) %>','<%= mid(ocombine.fitemlist(i).fYYYYMM,6,2) %>','01','<%= left(ocombine.fitemlist(i).fYYYYMM,4) %>','<%= mid(ocombine.fitemlist(i).fYYYYMM,6,2) %>','<%= LastDayOfThisMonth( left(ocombine.fitemlist(i).fYYYYMM,4),mid(ocombine.fitemlist(i).fYYYYMM,6,2)) %>','GNI');" onfocus="this.blur()">
		<%= FormatNumber(ocombine.fitemlist(i).fGNI,0) %>
		</a>
	</td>
	<td align="right">
		<a href="javascript:pop_detail_list('<%= left(ocombine.fitemlist(i).fYYYYMM,4) %>','<%= mid(ocombine.fitemlist(i).fYYYYMM,6,2) %>','01','<%= left(ocombine.fitemlist(i).fYYYYMM,4) %>','<%= mid(ocombine.fitemlist(i).fYYYYMM,6,2) %>','<%= LastDayOfThisMonth( left(ocombine.fitemlist(i).fYYYYMM,4),mid(ocombine.fitemlist(i).fYYYYMM,6,2)) %>','GNC');" onfocus="this.blur()">
		<%= FormatNumber(ocombine.fitemlist(i).fGNC,0) %>
		</a>
	</td>
	<td align="right">
		<a href="javascript:pop_detail_list('<%= left(ocombine.fitemlist(i).fYYYYMM,4) %>','<%= mid(ocombine.fitemlist(i).fYYYYMM,6,2) %>','01','<%= left(ocombine.fitemlist(i).fYYYYMM,4) %>','<%= mid(ocombine.fitemlist(i).fYYYYMM,6,2) %>','<%= LastDayOfThisMonth( left(ocombine.fitemlist(i).fYYYYMM,4),mid(ocombine.fitemlist(i).fYYYYMM,6,2)) %>','SFT');" onfocus="this.blur()">
		<%= FormatNumber(ocombine.fitemlist(i).fSFT,0) %>
		</a>
	</td>
	<td align="right">
		<a href="javascript:pop_detail_list('<%= left(ocombine.fitemlist(i).fYYYYMM,4) %>','<%= mid(ocombine.fitemlist(i).fYYYYMM,6,2) %>','01','<%= left(ocombine.fitemlist(i).fYYYYMM,4) %>','<%= mid(ocombine.fitemlist(i).fYYYYMM,6,2) %>','<%= LastDayOfThisMonth( left(ocombine.fitemlist(i).fYYYYMM,4),mid(ocombine.fitemlist(i).fYYYYMM,6,2)) %>','SPO');" onfocus="this.blur()">
		<%= FormatNumber(ocombine.fitemlist(i).fSPO,0) %>
		</a>
	</td>
    <td align="right">
		<a href="javascript:pop_detail_list('<%= left(ocombine.fitemlist(i).fYYYYMM,4) %>','<%= mid(ocombine.fitemlist(i).fYYYYMM,6,2) %>','01','<%= left(ocombine.fitemlist(i).fYYYYMM,4) %>','<%= mid(ocombine.fitemlist(i).fYYYYMM,6,2) %>','<%= LastDayOfThisMonth( left(ocombine.fitemlist(i).fYYYYMM,4),mid(ocombine.fitemlist(i).fYYYYMM,6,2)) %>','SPE');" onfocus="this.blur()">
		<%= FormatNumber(ocombine.fitemlist(i).fSPE,0) %>
		</a>
	</td>
	<td align="right">
		<a href="javascript:pop_detail_list('<%= left(ocombine.fitemlist(i).fYYYYMM,4) %>','<%= mid(ocombine.fitemlist(i).fYYYYMM,6,2) %>','01','<%= left(ocombine.fitemlist(i).fYYYYMM,4) %>','<%= mid(ocombine.fitemlist(i).fYYYYMM,6,2) %>','<%= LastDayOfThisMonth( left(ocombine.fitemlist(i).fYYYYMM,4),mid(ocombine.fitemlist(i).fYYYYMM,6,2)) %>','RTD');" onfocus="this.blur()">
		<%= FormatNumber(ocombine.fitemlist(i).fRTD,0) %>
		</a>
	</td>
	<td align="right">
	    <a href="javascript:pop_detail_list('<%= left(ocombine.fitemlist(i).fYYYYMM,4) %>','<%= mid(ocombine.fitemlist(i).fYYYYMM,6,2) %>','01','<%= left(ocombine.fitemlist(i).fYYYYMM,4) %>','<%= mid(ocombine.fitemlist(i).fYYYYMM,6,2) %>','<%= LastDayOfThisMonth( left(ocombine.fitemlist(i).fYYYYMM,4),mid(ocombine.fitemlist(i).fYYYYMM,6,2)) %>','XPR');" onfocus="this.blur()">
		<%= FormatNumber(ocombine.fitemlist(i).fXPR,0) %>
		</a>
	</td>

	<td align="right">
		<a href="javascript:pop_detail_list('<%= left(ocombine.fitemlist(i).fYYYYMM,4) %>','<%= mid(ocombine.fitemlist(i).fYYYYMM,6,2) %>','01','<%= left(ocombine.fitemlist(i).fYYYYMM,4) %>','<%= mid(ocombine.fitemlist(i).fYYYYMM,6,2) %>','<%= LastDayOfThisMonth( left(ocombine.fitemlist(i).fYYYYMM,4),mid(ocombine.fitemlist(i).fYYYYMM,6,2)) %>','ETC');" onfocus="this.blur()">
		<%= FormatNumber(ocombine.fitemlist(i).fETC,0) %>
		</a>
	</td>

	<td align="right">
		<b><%
		Select Case srcGbn
			Case "D"
				response.write FormatNumber((ocombine.fitemlist(i).fGNE + ocombine.fitemlist(i).fSPO + ocombine.fitemlist(i).fETC),0)
			Case Else
				response.write FormatNumber((ocombine.fitemlist(i).fSPO + ocombine.fitemlist(i).fETC),0)
		End Select
		%></b>
	</td>

	<% if (C_ADMIN_AUTH) then %><td>
	    <% if (DateAdd("m",+3,CDate(ocombine.fitemlist(i).fYYYYMM+"-01"))>now()) and (ocombine.fitemlist(i).fYYYYMM>="2014-01") then %>
	    <img src="/images/icon_reload.gif" onClick="refreshSmr('<%= ocombine.fitemlist(i).fYYYYMM %>');" style="cursor:pointer;">
	    <% end if %>
	</td><% end if %>
</tr>
<%
PPoint = ocombine.fitemlist(i).fpointsum
%>
<% next %>

<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
    <td height="25">�հ�</td>
    <td align="right"></td>
    <td align="right"><%= FormatNumber(totpointsum,0) %></td>
	<td align="right">
		<%= FormatNumber(totORD,0) %>
	</td>
	<% if (srcGbn="M") then %>
	<td align="right">
		<%= FormatNumber(totGOE,0) %>
	</td>
	<td align="right">
		<%= FormatNumber(totGPE,0) %>
	</td>
	<% else %>
	<td align="right">
		<%= FormatNumber(totGNE,0) %>
	</td>
	<% end if %>
	<td align="right">
		<%= FormatNumber(totGNI,0) %>
	</td>
	<td align="right">
		<%= FormatNumber(totGNC,0) %>
	</td>
	<td align="right">
		<%= FormatNumber(totSFT,0) %>
	</td>
	<td align="right">
		<%= FormatNumber(totSPO,0) %>
	</td>
    <td align="right">
		<%= FormatNumber(totSPE,0) %>
	</td>
	<td align="right">
		<%= FormatNumber(totRTD,0) %>
	</td>
    <td align="right">
		<%= FormatNumber(totXPR,0) %>
	</td>

	<td align="right">
		<%= FormatNumber(totETC,0) %>
	</td>
	<td align="right">
	</td>
	<% if (C_ADMIN_AUTH) then %><td></td><% end if %>
</tr>

<% else %>
<tr align="center" bgcolor="#FFFFFF">
	<td colspan="21">��ϵ� ������ �����ϴ�.</td>
</tr>
<% end if %>
</table>
<form name="frmAct" method="post" action="pointsum_process.asp">
<input type="hidden" name="mode">
<input type="hidden" name="yyyymm">
</form>
<%
Set ocombine = Nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
