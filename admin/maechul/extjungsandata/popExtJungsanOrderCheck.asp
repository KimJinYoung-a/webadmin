<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : ���޸� ���� Vs �ֹ�����
' Hieditor : 2018.04.22 �̻� ����
'			 2012.08.24 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/extjungsan/extjungsancls.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->

<%
Dim i
dim research : research = requestCheckvar(request("research"),10)
dim sellsite : sellsite = requestCheckvar(request("sellsite"),32)
dim page : page = requestCheckvar(request("page"),10)
dim difftp : difftp = requestCheckvar(request("difftp"),10)
dim chkdlvmonth : chkdlvmonth = requestCheckvar(request("chkdlvmonth"),10)
dim chkbysum : chkbysum = requestCheckvar(request("chkbysum"),10)

dim yyyy1, mm1, fromDate, toDate, dlvyyyy, dlvmm
yyyy1 = requestCheckvar(request("yyyy1"),4)
mm1 = requestCheckvar(request("mm1"),2)
dlvyyyy = requestCheckvar(request("dlvyyyy"),4)
dlvmm = requestCheckvar(request("dlvmm"),2)

if (yyyy1="") then yyyy1=LEFT(NOW(),4)
if (mm1="") then mm1=MID(NOW(),6,2)
if (dlvyyyy="") then dlvyyyy=yyyy1
if (dlvmm="") then dlvmm=mm1
if (page="") then page=1
if (difftp="") then difftp="2"  ''���ǸŰ�.

fromDate = yyyy1+"-"+mm1+"-01"
toDate = dateADD("m",1,fromDate)

dim oCExtJungsanDiff
SET oCExtJungsanDiff = new CExtJungsan
oCExtJungsanDiff.FPageSize = 2000
oCExtJungsanDiff.FCurrPage = page
oCExtJungsanDiff.FRectSellSite = sellsite
oCExtJungsanDiff.FRectJungsanType = "C"
oCExtJungsanDiff.FRectStartdate = fromDate
oCExtJungsanDiff.FRectEndDate = toDate
oCExtJungsanDiff.FRectDiffType = difftp
if (chkdlvmonth<>"") then
    oCExtJungsanDiff.FRectDlvMonth = dlvyyyy+"-"+dlvmm
end if
oCExtJungsanDiff.FRectCheckBySum = chkbysum
oCExtJungsanDiff.getExtJungsanOrderDiffList_replica

dim FormatDotNo : FormatDotNo=0
%>
<script language='javascript'>
function popByExtorderserial(iextorderserial){
	var iUrl = "/admin/maechul/extjungsandata/extJungsanMapEdit.asp?menupos=<%=menupos%>&page=1&research=on";
	iUrl += "&sellsite=<%=sellsite%>"
	iUrl += "&searchfield=extOrderserial&searchtext="+iextorderserial;
	var popwin = window.open(iUrl,"extJungsanMapEdit","width=1400,height=800,scrollbars=yes,resizable=yes,status=yes");

	popwin.focus();

}

function actCpnMapByorderserial(sellsite,extOrderserial,Orderserial){
	var popwin = window.open("","extJungsanEditProc","width=600,height=300");

	popwin.location.href="/admin/maechul/extjungsandata/extJungsan_process.asp?sellsite="+sellsite+"&extOrderserial="+extOrderserial+"&Orderserial="+Orderserial+"&mode=mapcpnbyorderserial";

	popwin.focus();
}
</script>
<!-- �˻� ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<form name="frm" method="get" action="">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<input type="hidden" name="research" value="on">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		* ���޸�:
		<%= getJungsanXsiteComboHTML("sellsite",sellsite,"") %>

		&nbsp;

		* �����:
		<% DrawYMBox yyyy1,mm1 %>
        &nbsp;

        * ����Ÿ��
        <select class="select" name="difftp">
        <option value="0" <%=CHKIIF(difftp="0","selected","") %> >��ü
        <option value="1" <%=CHKIIF(difftp="1","selected","") %> >�ǸŰ�
        <option value="2" <%=CHKIIF(difftp="2","selected","") %> >���ǸŰ�
        <option value="3" <%=CHKIIF(difftp="3","selected","") %> >����
        </select>

        &nbsp;
        * <input type="checkbox" name="chkdlvmonth" <%=CHKIIF(chkdlvmonth<>"","checked","")%> >�ֹ�����
        <% DrawYMBoxdynamic "dlvyyyy",dlvyyyy,"dlvmm",dlvmm,"" %>

		&nbsp;
        * <input type="checkbox" name="chkbysum" <%=CHKIIF(chkbysum<>"","checked","")%> >�հ�κ���
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" style="width:70px;height:50px;" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="left" bgcolor="#FFFFFF" >
	<td>
	* �����ݾ��� �������� ���� �� �ݿ��Ǵ� ���޸� : SSG, Hmall, WMP, wmpfashion, LotteiMall, LotteCom, LFMall, coupang
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->
<p  >
<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="22">
		�˻���� : <b><%= oCExtJungsanDiff.FTotalcount %></b>
		&nbsp;
		<% if oCExtJungsanDiff.FTotalcount>=oCExtJungsanDiff.FPageSize then %>
        (�ִ� <%=FormatNumber(oCExtJungsanDiff.FPageSize,0)%> ��)
        <% end if %>
	</td>
</tr>
<form name="frm1" method="post">
<input type="hidden" name="mode" value="">
<input type="hidden" name="xSiteId" value="">
<input type="hidden" name="idx" value="">
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="80">��������</td>
	<td width="150">����<br>�ֹ���ȣ</td>
	<td width="60">����<br>�ֹ�����</td>
	<td width="80">����<br>���ֹ���ȣ</td>
	<td width="40">����</td>

	<td width="60">�ǸŰ�</td>
	<td width="60">���޺δ�<br>����</td>
	<td width="60">���ٺδ�<br>����</td>
	<td width="60">������</td>
	<td width="80">���ֹ���ȣ</td>
	<td width="100">��ǰ�ڵ�</td>
	<td width="60">�ɼ��ڵ�</td>
    <td width="60">�ֹ�<br>�ǸŰ�</td>
    <td width="60">�ֹ�<br>���ǸŰ�</td>
    <td width="60">�ֹ�<br>����</td>
    <td width="60">�ֹ�<br>�����</td>
	<td width="60">�ֹ�<br>�����</td>
	<td width="60">�ֹ�<br>������</td>
	<td>���</td>

    <td width="60">�ǸŰ���</td>
    <td width="60">���ǸŰ���</td>
    <td width="60">������</td>
</tr>

<% if oCExtJungsanDiff.FresultCount<1 then %>
<tr align="center" bgcolor="FFFFFF" onmouseover=this.style.background="F1F1F1"; onmouseout=this.style.background="FFFFFF";>
    <td colspan="22">
        <% if (sellsite="") then %>
        [���� ���޸��� ���� �ϼ���.]
        <% else %>
        [�˻������ �����ϴ�.]
        <% end if %>
    </td>
</tr>
<% else %>
<% for i=0 to oCExtJungsanDiff.FresultCount -1 %>
<tr align="center" bgcolor="FFFFFF" onmouseover=this.style.background="F1F1F1"; onmouseout=this.style.background="FFFFFF";>
	<td><%= oCExtJungsanDiff.FItemList(i).FextMeachulDate %></td>
	<td><a href="#" onClick="popByExtorderserial('<%= oCExtJungsanDiff.FItemList(i).FextOrderserial %>');return false;"><%= oCExtJungsanDiff.FItemList(i).FextOrderserial %></a></td>
	<td><%= oCExtJungsanDiff.FItemList(i).FextOrderserSeq %></td>
	<td>
		<% if Null2Blank(oCExtJungsanDiff.FItemList(i).FextOrgOrderserial)<>"" then %>
		<a href="#" onClick="popByExtorderserial('<%= oCExtJungsanDiff.FItemList(i).FextOrgOrderserial %>');return false;"><%= oCExtJungsanDiff.FItemList(i).FextOrgOrderserial %></a>
		<% end if %>
	</td>
	<td><%= oCExtJungsanDiff.FItemList(i).FextItemNo %></td>
	<td align="right"><%= FormatNumber(oCExtJungsanDiff.FItemList(i).FextItemCost, FormatDotNo) %></td>
	<td align="right"><%= FormatNumber(oCExtJungsanDiff.FItemList(i).FextOwnCouponPrice, FormatDotNo) %></td>
	<td align="right"><%= FormatNumber(oCExtJungsanDiff.FItemList(i).FextTenCouponPrice, FormatDotNo) %></td>
	<td align="right"><%= FormatNumber(oCExtJungsanDiff.FItemList(i).FextReducedPrice, FormatDotNo) %></td>

	<td><%= oCExtJungsanDiff.FItemList(i).FOrgOrderserial %></td>
	<td><%= oCExtJungsanDiff.FItemList(i).Fitemid %></td>
	<td><%= oCExtJungsanDiff.FItemList(i).Fitemoption %></td>
    <td><% IF NOT isNULL(oCExtJungsanDiff.FItemList(i).Forderitemcost) then response.write FormatNumber(oCExtJungsanDiff.FItemList(i).Forderitemcost,0) END IF %></td>
    <td><% IF NOT isNULL(oCExtJungsanDiff.FItemList(i).Forderreducedprice) then response.write FormatNumber(oCExtJungsanDiff.FItemList(i).Forderreducedprice,0) END IF %></td>
    <td><% IF NOT isNULL(oCExtJungsanDiff.FItemList(i).Forderitemno) then response.write FormatNumber(oCExtJungsanDiff.FItemList(i).Forderitemno,0) END IF %></td>
    <td>
		<% if isNULL(oCExtJungsanDiff.FItemList(i).Forderbeasongdate) then %>

		<% elseif (LEFT(oCExtJungsanDiff.FItemList(i).Forderbeasongdate,7)<>yyyy1&"-"&mm1) then %>
		<font color="#CCCCCC"><%= oCExtJungsanDiff.FItemList(i).Forderbeasongdate %></font>
		<% else %>
		<%= oCExtJungsanDiff.FItemList(i).Forderbeasongdate %>
		<% end if %>
	</td>
	<td>
		<% if isNULL(oCExtJungsanDiff.FItemList(i).Fdlvfinishdt) then %>

		<% elseif (LEFT(oCExtJungsanDiff.FItemList(i).Fdlvfinishdt,7)<>yyyy1&"-"&mm1) then %>
		<font color="#CCCCCC"><%= oCExtJungsanDiff.FItemList(i).Fdlvfinishdt %></font>
		<% else %>
		<%= oCExtJungsanDiff.FItemList(i).Fdlvfinishdt %>
		<% end if %>
	</td>
	<td>
		<% if isNULL(oCExtJungsanDiff.FItemList(i).Fjungsanfixdate) then %>

		<% elseif (LEFT(oCExtJungsanDiff.FItemList(i).Fjungsanfixdate,7)<>yyyy1&"-"&mm1) then %>
		<font color="#CCCCCC"><%= oCExtJungsanDiff.FItemList(i).Fjungsanfixdate %></font>
		<% else %>
		<%= oCExtJungsanDiff.FItemList(i).Fjungsanfixdate %>
		<% end if %>

	</td>
	<td>
		<%=oCExtJungsanDiff.FItemList(i).getBigoStr%>

		<% if ((sellsite="ssg") or (sellsite="hmall1010") or (sellsite="WMP") or (sellsite="wmpfashion") or (sellsite="lotteon")) and (chkbysum="") then %>

		<% else %>
		<% if (oCExtJungsanDiff.FItemList(i).isCpnValEditAvailRow) then %>
			<% if (oCExtJungsanDiff.FItemList(i).getBigoStr<>"") then %><br><% end if %>
			<input type="button" value="�������ݿ�" onClick="actCpnMapByorderserial('<%=oCExtJungsanDiff.FItemList(i).Fsellsite%>','<%=CHKIIF(NULL2Blank(oCExtJungsanDiff.FItemList(i).FextOrgOrderserial)<>"",oCExtJungsanDiff.FItemList(i).FextOrgOrderserial,oCExtJungsanDiff.FItemList(i).FextOrderserial)%>','<%= oCExtJungsanDiff.FItemList(i).FOrgOrderserial %>')">
		<% end if %>
		<% end if %>
	</td>
	<td><%= FormatNumber(oCExtJungsanDiff.FItemList(i).getJOdiffItemCost,FormatDotNo) %></td>
    <td><%= FormatNumber(oCExtJungsanDiff.FItemList(i).getJOdiffReducedprice,FormatDotNo) %></td>
    <td><%= FormatNumber(oCExtJungsanDiff.FItemList(i).getJOdiffitemno,FormatDotNo) %></td>
</tr>
<% next %>
<% end if %>

<tr height="25" bgcolor="FFFFFF">
	<td colspan="22" align="center">
		<% if oCExtJungsanDiff.HasPreScroll then %>
		<a href="javascript:NextPage('<%= oCExtJungsanDiff.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + oCExtJungsanDiff.StartScrollPage to oCExtJungsanDiff.FScrollCount + oCExtJungsanDiff.StartScrollPage - 1 %>
			<% if i>oCExtJungsanDiff.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oCExtJungsanDiff.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	</td>
</tr>
</form>
</table>

<%
set oCExtJungsanDiff = Nothing
%>


<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
