<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdminNoCache.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/jungsan/jungsanTaxCls.asp"-->
<%
dim makerid, yyyy1,mm1, jgubun, targetGbn, groupid, page, finishflag, taxtype,jacctcd,noincTen
makerid = requestCheckvar(request("makerid"),32)
yyyy1   = requestCheckvar(request("yyyy1"),10)
mm1     = requestCheckvar(request("mm1"),10)
jgubun  = requestCheckvar(request("jgubun"),10)
targetGbn	= requestCheckvar(request("targetGbn"),10)
groupid  	= requestCheckvar(request("groupid"),10)
finishflag 	= requestCheckvar(request("finishflag"),10)
taxtype   	= requestCheckvar(request("taxtype"),10)
jacctcd    	= requestCheckvar(request("jacctcd"),10)
noincTen    = requestCheckvar(request("noincTen"),10)
page    	= requestCheckvar(request("page"),10)
if (page="") then page=1

if (yyyy1="") then
    yyyy1 = LEFT(dateadd("m",-1,now()),4)
    mm1 = MID(dateadd("m",-1,now()),6,2)
end if

if (jgubun="") then
    jgubun = "MM"
end if

dim ojungsanTax
set ojungsanTax = new CUpcheJungsanTax
ojungsanTax.FPageSize = 2500
ojungsanTax.FCurrPage = page
ojungsanTax.FRectMakerid = makerid
ojungsanTax.FRectYYYYMM = yyyy1+"-"+mm1
ojungsanTax.FRectJGubun = jgubun
ojungsanTax.FRectTargetGbn = targetGbn
ojungsanTax.FRectGroupid = groupid
ojungsanTax.FRectFinishFlag = finishflag
ojungsanTax.FRectTaxType = taxtype
ojungsanTax.FRectjacctcd = jacctcd
ojungsanTax.FRectNotIncTen = noincTen
ojungsanTax.getMonthUpcheJungsanListAdmAll


dim i
%>

<!-- �������Ϸ� ���� ��� �κ� -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
Response.ContentType = "application/unknown"
Response.Write("<meta http-equiv='Content-Type' content='text/html; charset=EUC-KR'>")

Response.ContentType = "application/vnd.ms-excel"
Response.ContentType = "application/x-msexcel"
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition", "attachment;filename="&yyyy1&"-"&mm1&CHKIIF(jgubun="CC","������","����")&"���곻��.xls"
%>
<style type="text/css">
/* ���� �ٿ�ε�� ����� ���ڷ� ǥ�õ� ��� ���� */
.txt {mso-number-format:'\@'}
</style>
</head>
<body>


<% if (jgubun="CC") then %>
<table width="100%" border="0" align="center" class="a" cellpadding="4" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<table width="100%" border="0" align="center" class="a" cellpadding="4" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF">
    <td colspan="19" align="left"><strong>* ������ ���� ����</strong> <font color=red>(������ ���� ��꼭�� <b>�ٹ�����</b>���� <b>�ϰ� ����</b>�մϴ�.)</font></td>
    <td colspan="2" align="right">�� <%=ojungsanTax.FTotalcount%> �� <%=page%> / <%=ojungsanTax.FTotalpage%></td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="26">
    <td width="60" class="txt">�����</td>
    <td width="60" class="txt">����ó</td>
    <td width="50" class="txt">����<br>����</td>
    <td width="50" class="txt">�׷��ڵ�</td>
    <td width="50" class="txt">ERPCode</td>
    <td width="90" >ȸ���</td>
    <td width="90" >����ڹ�ȣ</td>
    <td width="90" class="txt">�귣��ID</td>
    <td width="180"class="txt">���곻��</td>
    <td width="90" >��������</td>
    <td width="90" >�����Ǹűݾ�<br>(���»� �����)</td>
    <td width="80" >��ۺ�<br>(�Ǹűݾ�)</td>
    <td width="80" >��ǰ�Ǹ�<br>������</td>
    <td width="80" >��������<br>������</td>
    <td width="100">���޿�����<br>(��ǰ)</td>
    <td width="80">���޿�����<br>(��ۺ�)</td>
  	<td width="80">�߰������<br>(�߰���ۺ�)</td>
  	<td width="80">�߰������<br>(��ǰ��ۺ��)</td>
  	<td width="80">�߰������<br>(��Ÿ���θ�ǵ�)</td>
  	<td width="80">���޿�����</td>
    <td width="90" >��꼭����</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="26">
    <td>�հ�</td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td align="left"></td>
    <td align="left"></td>
    <td align="left"></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FSumaryOneItem.FPrdMeachulsum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FSumaryOneItem.FdlvMeachulsum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FSumaryOneItem.FPrdCommissionSum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FSumaryOneItem.FPgCommissionSum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FSumaryOneItem.FprdJungsanSum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FSumaryOneItem.getOriginDlvJungsanSum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FSumaryOneItem.getAddDlvJungsanSum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FSumaryOneItem.getEtcDlvJungsanSum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FSumaryOneItem.getPromotionJungsanSum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FSumaryOneItem.getToTalJungsanSum,0)%></td>
    <td></td>
</tr>
<% for i=0 to ojungsanTax.FresultCount-1 %>
<tr align="center" bgcolor="#FFFFFF" height="30">
    <td class="txt"><%=ojungsanTax.FItemList(i).Fyyyymm%></td>
    <td class="txt"><%=ojungsanTax.FItemList(i).getTargetNm%></td>
    <td class="txt"><%=ojungsanTax.FItemList(i).getItemVatTypeName%></td>
    <td class="txt"><%=ojungsanTax.FItemList(i).Fgroupid%></td>
    <td class="txt"><%=ojungsanTax.FItemList(i).FerpCust_cd%></td>
    <td class="txt"><%=ojungsanTax.FItemList(i).Fcompany_name%></td>
    <td class="txt"><%=ojungsanTax.FItemList(i).Fcompany_no%></td>
    <td align="left"><%=ojungsanTax.FItemList(i).Fmakerid%></td>
    <td align="left"><%=ojungsanTax.FItemList(i).Ftitle%></td>
    <td align="left"><%=ojungsanTax.FItemList(i).Facc_nm%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FItemList(i).FPrdMeachulsum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FItemList(i).FdlvMeachulsum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FItemList(i).FPrdCommissionSum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FItemList(i).FPgCommissionSum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FItemList(i).FprdJungsanSum,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FItemList(i).getOriginDlvJungsanSum,0)%></td>
    <td align="right">
    <% if isNULL(ojungsanTax.FItemList(i).getAddDlvJungsanSum) then %>

    <% else %>
    <%=FormatNumber(ojungsanTax.FItemList(i).getAddDlvJungsanSum,0)%>
    <% end if %>
    </td>
    <td align="right">
    <% if isNULL(ojungsanTax.FItemList(i).getEtcDlvJungsanSum) then %>
    <% else %>
    <%=FormatNumber(ojungsanTax.FItemList(i).getEtcDlvJungsanSum,0)%>
    <% end if%>
    </td>
    <td align="right">
    <% if isNULL(ojungsanTax.FItemList(i).getPromotionJungsanSum) then %>
    <% else %>
    <%=FormatNumber(ojungsanTax.FItemList(i).getPromotionJungsanSum,0)%>
    <% end if%>
    </td>
    <td align="right"><%=FormatNumber(ojungsanTax.FItemList(i).getToTalJungsanSum,0)%></td>
    <td><%=ojungsanTax.FItemList(i).GetTaxEvalStateName%></td>

</tr>
<% next %>


<% else %>

<table width="100%" border="0" align="center" class="a" cellpadding="4" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF">
    <td colspan="14" align="left"><strong>* ���� ���� ����</strong> (���»翡�� �ٹ��������� ������ �ּž� �մϴ�.) (�Ե����� �Ǹ� ���� �� ������ �Ǹ� ������ ������������ ó�� �˴ϴ�.)</td>
    <td colspan="2" align="right">�� <%=ojungsanTax.FTotalcount%> �� <%=page%> / <%=ojungsanTax.FTotalpage%></td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="26">
    <td width="60" class="txt">�����</td>
    <td width="60" class="txt">����ó</td>
    <td width="50" class="txt">����<br>����</td>
    <td width="50" class="txt">�׷��ڵ�</td>
    <td width="50" class="txt">ERPCode</td>
   <td width="90" >ȸ���</td>
   <td width="90" >����ڹ�ȣ</td>
    <td width="90" class="txt">�귣��ID</td>
    <td width="170" >���곻��</td>
    <td width="90" >��������</td>
    <td width="90" >�԰��и���<br>(��ǰ���޾�)</td>
    <td width="90" >�Ǹźи���<br>(��ǰ���޾�)</td>
    <td width="90" >��Ÿ����<br>(����)</td>
    <td width="90" >��Ÿ����<br>(��ۺ�)</td>
    <td width="90" >��Ÿ�������<br>(��Ÿ���)</td>
    <td width="90" >���޿�����<br>(���»�����)</td>
    <td width="90" >��꼭����</td>
</tr>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>" height="30">
    <td>�հ�</td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td align="left"></td>
    <td align="left"></td>
    <td align="left"></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FSumaryOneItem.FMSuply,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FSumaryOneItem.FSSuply,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FSumaryOneItem.FESuply,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FSumaryOneItem.FDSuply,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FSumaryOneItem.FCSuply,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FSumaryOneItem.getToTalJungsanSum,0)%></td>
    <td></td>
</tr>
<% for i=0 to ojungsanTax.FresultCount-1 %>
<tr align="center" bgcolor="#FFFFFF" height="30">
    <td class="txt"><%=ojungsanTax.FItemList(i).Fyyyymm%></td>
    <td class="txt"><%=ojungsanTax.FItemList(i).getTargetNm%></td>
    <td class="txt"><%=ojungsanTax.FItemList(i).getTaxtypeName%></td>
    <td class="txt"><%=ojungsanTax.FItemList(i).Fgroupid%></td>
    <td class="txt"><%=ojungsanTax.FItemList(i).FerpCust_cd%></td>
    <td class="txt"><%=ojungsanTax.FItemList(i).Fcompany_name%></td>
    <td class="txt"><%=ojungsanTax.FItemList(i).Fcompany_no%></td>
    <td align="left"><%=ojungsanTax.FItemList(i).Fmakerid%></td>
    <td align="left"><%=ojungsanTax.FItemList(i).Ftitle%></td>
    <td align="left"><%=ojungsanTax.FItemList(i).Facc_nm%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FItemList(i).FMSuply,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FItemList(i).FSSuply,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FItemList(i).FESuply,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FItemList(i).FDSuply,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FItemList(i).FCSuply,0)%></td>
    <td align="right"><%=FormatNumber(ojungsanTax.FItemList(i).getToTalJungsanSum,0)%></td>
    <td><%=ojungsanTax.FItemList(i).GetTaxEvalStateName%></td>

</tr>
<% next %>

</table>
<% end if %>
</body>
</html>
<%
set ojungsanTax = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->