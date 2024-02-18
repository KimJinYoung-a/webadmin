<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �¶��� �������� ���ϸ��� & ��ġ�� ���հ���
' History : 2013.11.12 �ѿ�� ����
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

dim dateGubun
Dim i, yyyy1,mm1, ocombine, srcGbn, targetGbn, GbnCd, page
	yyyy1   = requestcheckvar(request("yyyy1"),10)
	mm1     = requestcheckvar(request("mm1"),10)

	srcGbn     	= requestcheckvar(request("srcGbn"),1)
	targetGbn   = requestcheckvar(request("targetGbn"),4)
	GbnCd     	= requestcheckvar(request("GbnCd"),16)
    page  		= requestcheckvar(request("page"),10)
	dateGubun   = requestcheckvar(request("dateGubun"),1)

if (yyyy1="") then yyyy1 = Cstr(Year( dateadd("m",-3,date()) ))
if (mm1="") then mm1 = Cstr(Month( dateadd("m",-3,date()) ))
if (page="") then page=1
if (dateGubun="") then dateGubun="M"

Set ocombine = New ccombine_point_deposit

	if (dateGubun = "M") then
		ocombine.FRectYYYYMM = yyyy1+"-"+mm1
	else
		ocombine.FRectYYYYMM = yyyy1
	end if

	ocombine.FRectsrcGbn = srcGbn
	ocombine.FRecttargetGbn = targetGbn
	ocombine.FRectGbnCd = GbnCd
	ocombine.FPageSize = 2000
	ocombine.FCurrPage	= page

	if (targetGbn<>"") then
	ocombine.fcombine_point_deposit_Detail_list()
    end if

%>

<script language="javascript">

function searchSubmit(){
	frm.submit();
}


function NextPage(page){
    var frm = document.frm;
	frm.page.value = page;
	frm.submit();
}

function popModifyDate(idx, yyyymmdd, srcGbn) {
    <% if C_ADMIN_AUTH then %>
    // �ϴ� ��ġ�ݸ�
    alert('������ ����');
	var popwin = window.open("popModifyDate.asp?idx=" + idx + "&yyyymmdd=" + yyyymmdd + "&srcGbn=" + srcGbn,"popModifyDate","width=200 height=120 scrollbars=yes resizable=yes status=yes");
	popwin.focus();
    <% else %>
    alert('�����ڸ� ���������մϴ�.');
    <% end if %>
}

</script>

<!-- �˻� ���� -->
<form name="frm" method="get" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td width="70" bgcolor="<%= adminColor("gray") %>">�˻�</td>
	<td align="left">
		<table class="a">
		<tr>
			<td height="25">
				* ��¥ : <% DrawYMBoxdynamic "yyyy1",yyyy1,"mm1",mm1,"" %>
				* ��¥���� :
				<input type="radio" name="dateGubun" value="M" <% if (dateGubun = "M") then %>checked<% end if %>> ����
				<input type="radio" name="dateGubun" value="Y" <% if (dateGubun = "Y") then %>checked<% end if %>> ������
				<p>
				* ���� : <% drawoffshop_commoncode "srcGbn", srcGbn, "srcGbn", "MAIN", "", " onchange='searchSubmit()'" %>
				&nbsp;&nbsp;
				* ä�� : <% drawoffshop_commoncode "targetGbn", targetGbn, "targetGbn", "MAIN", "", " onchange='searchSubmit()'" %>
				&nbsp;&nbsp;
				* ���� : <% drawoffshop_commoncode "GbnCd", GbnCd, "GbnCd", "MAIN", srcGbn, " onchange='searchSubmit()'" %>
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

* ���� ������ ���ԵǾ� �ֽ��ϴ�.<br />
* ���Ӹ����� ���� ������ ���ܵ˴ϴ�.

<p />

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="20">
		�˻���� : �� <%= ocombine.FTotalCount %> �� | <%= FormatNumber(ocombine.FTotalSum,0) %> point | page <%=page%>/<%=ocombine.FTotalPage%> | ����page�˻��� :<%= ocombine.FresultCount %>
	</td>
</tr>
<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
    <td>����Ʈ</td>
    <td>�ֹ���ȣ</td>
    <td>�ֹ���ȣ2</td>
    <td>USERID</td>
    <td>��¥</td>
    <td>����</td>
    <td>�ݾ�</td>
    <% if srcGbn="M" and GbnCd="GNE" then %>
    <td>���</td>
    <% end if %>
	<% if srcGbn="G" and GbnCd="GNE" then %>
	<td>���ID</td>
	<% end if %>
</tr>

<%
dim totpointsum
	totpointsum = 0
dim isOffORDMile
isOffORDMile = (srcGbn="M") and (targetGbn="OF") and (GbnCd="ORD")

if ocombine.FresultCount > 0 then

For i = 0 To ocombine.FresultCount -1

totpointsum = totpointsum + ocombine.fitemlist(i).FiPoint
%>
<% if srcGbn="G" then %>
	<tr bgcolor="#FFFFFF" align="center" onmouseover=this.style.background="#F1F1F1"; onmouseout=this.style.background='#FFFFFF';>
<% else %>
	<% if (isOffORDMile) and ((LEFT(ocombine.fitemlist(i).Forderserial,4)<replace(MID(ocombine.fitemlist(i).Fyyyymmdd,3,5),"-","")) or DateDiff("d", ocombine.fitemlist(i).GetYYYYMMDD(), ocombine.fitemlist(i).Fyyyymmdd) >= 7) then %>
	<tr bgcolor="#FFAAAA" align="center" onmouseover=this.style.background="#F1F1F1"; onmouseout=this.style.background='#FFAAAA';>
	<% elseif (srcGbn="G") and (GbnCd="ORD") and not isNULL(ocombine.fitemlist(i).Fcanceldate)  then %>
		<tr bgcolor="#FFAAAA" align="center" onmouseover=this.style.background="#F1F1F1"; onmouseout=this.style.background='#FFAAAA';>
	<% else %>
	<tr bgcolor="#FFFFFF" align="center" onmouseover=this.style.background="#F1F1F1"; onmouseout=this.style.background='#FFFFFF';>
	<% end if %>
<% end if %>
    <td height="25">
		<%= ocombine.fitemlist(i).FtargetGbn %>
	</td>
	<td><%= ocombine.fitemlist(i).Forderserial %></td>
	<td><%= ocombine.fitemlist(i).FsubOrderserial %></td>
	<td><%= ocombine.fitemlist(i).Fuserid %></td>
	<td>
	    <% if (srcGbn="G") and (GbnCd="ORD") and not isNULL(ocombine.fitemlist(i).Fcanceldate)  then %>
	    <%= ocombine.fitemlist(i).Fcanceldate %>
	    <% else %>
        	<% if ((srcGbn="D") or (srcGbn="G")) and (ocombine.fitemlist(i).FtargetGbn="ON") and (GbnCd="SPE") then %>
	    		<a href="javascript:popModifyDate(<%= ocombine.fitemlist(i).Fidx %>, '<%= ocombine.fitemlist(i).Fyyyymmdd %>', '<%= srcGbn %>')"><%= ocombine.fitemlist(i).Fyyyymmdd %></a>
        	<% else %>
        		<%= ocombine.fitemlist(i).Fyyyymmdd %>
        	<% end if %>
	    <% end if %>
	</td>
	<td><%= ocombine.fitemlist(i).FDtlDesc %></td>
	<td align="right"><%= FormatNumber(ocombine.fitemlist(i).FiPoint,0) %></td>
    <% if srcGbn="M" and GbnCd="GNE" then %>
    <td><%= chkIIF(ocombine.fitemlist(i).Fipkumdiv="1100","��������","�̺�Ʈ����") %></td>
    <% end if %>
    <% if srcGbn="G" and GbnCd="GNE" then %>
    <td><%= ocombine.fitemlist(i).FregUserid %></td>
    <% end if %>
</tr>
<% next %>

<tr bgcolor="<%= adminColor("tabletop") %>" align="center">
    <td height="25" colspan=6>�հ�</td>
    <td align="right"><%= FormatNumber(totpointsum,0) %></td>
    <% if (srcGbn="M" or srcGbn="G") and GbnCd="GNE" then %>
    <td></td>
    <% end if %>
</tr>
<tr bgcolor="#FFFFFF">
    <td colspan="<%=chkIIF((srcGbn="M" or srcGbn="G") and GbnCd="GNE","8","7")%>" align="center">
        <% if ocombine.HasPreScroll then %>
		<a href="javascript:NextPage('<%= ocombine.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + ocombine.StartScrollPage to ocombine.FScrollCount + ocombine.StartScrollPage - 1 %>
			<% if i>ocombine.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if ocombine.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
    </td>
</tr>
<% else %>
<tr align="center" bgcolor="#FFFFFF">
	<td colspan="20">
	<% if (targetGbn="") then %>
	<strong>ä�� ����</strong>�� ���� �����ϼ���
	<% else %>
	�˻� �����  �����ϴ�.
	<% end if %>
	</td>
</tr>
<% end if %>
</table>

<%
Set ocombine = Nothing
%>
<!-- #include virtual="/common/lib/commonbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
