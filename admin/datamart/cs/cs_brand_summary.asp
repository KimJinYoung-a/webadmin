<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  �귣�� CS ��� (��ǰ, ǰ����)
' History : 2009.12.19 ������ ����
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/datamart/brandcssummaryclass.asp"-->
<%
dim makerid, yyyy1, mm1
dim ck_date, research, page
dim stdate, cdl
dim isupchebeasong,divcd, casegroup

makerid     = RequestCheckVar(request("makerid"),32)
yyyy1       = RequestCheckVar(request("yyyy1"),32)
mm1         = RequestCheckVar(request("mm1"),32)
ck_date     = RequestCheckVar(request("ck_date"),32)
research    = RequestCheckVar(request("research"),32)
page        = RequestCheckVar(request("page"),9)
divcd       = RequestCheckVar(request("divcd"),4)
isupchebeasong = RequestCheckVar(request("isupchebeasong"),4)
casegroup   = RequestCheckVar(request("casegroup"),4)
cdl   = RequestCheckVar(request("cdl"),3)

if yyyy1="" then
	stdate = CStr(Now)
	stdate = DateSerial(Left(stdate,4), CLng(Mid(stdate,6,2)),1)
	yyyy1 = Left(stdate,4)
	mm1 = Mid(stdate,6,2)
end if

if (research<>"on") and (ck_date="") then ck_date="on"
if (research<>"on") and (casegroup="") then casegroup="on"
if (page="") then page=1

dim i

dim obrandCs
set obrandCs = new CBrandCSSummary
obrandCs.FPageSize = 100
obrandCs.FCurrPage = page
if (ck_date="on") then
    obrandCs.FRectYYYYMM = yyyy1 + "-" + mm1
end if

obrandCs.FRectNotIncludeETC     = "on"
obrandCs.FRectIsupchebeasong    = isupchebeasong
obrandCs.FRectDivCd             = divcd
obrandCs.FRectCDL               = cdl
obrandCs.FRectMakerid           = makerid

if (makerid<>"") or (ck_date="on") then

    if (casegroup="on") then
        obrandCs.getBrandCsSummary_GubunGroupNew
    else
        obrandCs.getBrandCssummary
    end if
end if

dim nYYYYMMDD
%>
<script language='javascript'>
function goPage(page){
    document.frm.page.value = page;
    document.frm.submit();
}

function popBrandCSAsDetailList(yyyymmdd1,yyyymmdd2,imakerid,idivcd,igubunNm){
    var iisupchebeasong = '<%= isupchebeasong %>';
    var ick_date = 'on';
    var iyyyy1 = yyyymmdd1.substr(0,4);
    var imm1   = yyyymmdd1.substr(5,2);
    var idd1   = yyyymmdd1.substr(8,2);
    var iyyyy2 = yyyymmdd2.substr(0,4);
    var imm2   = yyyymmdd2.substr(5,2);
    var idd2   = yyyymmdd2.substr(8,2);

    var popwin = window.open('/admin/csreport/brandcsdetail.asp?makerid=' + imakerid + '&divcd=' + idivcd + '&isupchebeasong=' + iisupchebeasong + '&gubunNm=' + igubunNm + '&ck_date=' + ick_date + '&yyyy1=' + iyyyy1 + '&mm1=' + imm1 + '&dd1=' + idd1 + '&yyyy2=' + iyyyy2 + '&mm2=' + imm2 + '&dd2=' + idd2,'popBrandCSAsDetailList','width=900,height=700,scrollbars=yes,resizable=yes');
    popwin.focus();
}

function popBrandCSSum(yyyymmdd1,yyyymmdd2,imakerid,idivcd,igubunNm){
    var iisupchebeasong = "";
    var ick_date = "on";
    var iyyyy1 = yyyymmdd1.substr(0,4);
    var imm1   = yyyymmdd1.substr(5,2);
    var idd1   = yyyymmdd1.substr(8,2);
    var iyyyy2 = yyyymmdd2.substr(0,4);
    var imm2   = yyyymmdd2.substr(5,2);
    var idd2   = yyyymmdd2.substr(8,2);

    var popwin = window.open("/admin/csreport/brandcs_sum.asp?makerid=" + imakerid + "&divcd=" + idivcd + "&isupchebeasong=" + iisupchebeasong + "&gubunNm=" + igubunNm + "&ck_date=" + ick_date + "&yyyy1=" + iyyyy1 + "&mm1=" + imm1 + "&dd1=" + idd1 + "&yyyy2=" + iyyyy2 + "&mm2=" + imm2 + "&dd2=" + idd2,"popBrandCSSum","width=900,height=700,scrollbars=yes,resizable=yes");
    popwin.focus();
}

</script>
<!-- ǥ ��ܹ� ����-->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get"   >
	<input type="hidden" name="menupos" value="<%=menupos%>">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="research" value="on">
  	<tr align="center" bgcolor="#FFFFFF" >
		<td width="50" bgcolor="<%= adminColor("gray") %>" rowspan="2">�˻�<br>����</td>
		<td align="left">
			<table border="0" width="100%" cellpadding="3" cellspacing="0" class="a">
			<tr>
				<td>
    				<input type="checkbox" name="ck_date" <%= ChkIIF(ck_date="on","checked","") %> onClick="chkComp(this.checked);">�ش�� : <% DrawYMBox yyyy1,mm1 %>(ó���Ϸ���)
    				&nbsp;&nbsp;
    				�귣��: <% drawSelectBoxDesignerwithName "makerid", makerid %>
    				&nbsp;&nbsp;
    				ī�װ�: <% SelectBoxCategoryLarge cdl %>

				</td>
			</tr>
			</table>
        </td>
		<td  width="50" bgcolor="<%= adminColor("gray") %>" rowspan="2">
		    <input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
	    </td>
	</tr>
	<tr bgcolor="#FFFFFF">
	    <td align="left">
	        <table border="0" width="100%" cellpadding="3" cellspacing="0" class="a">
			<tr>
				<td>
				    ��۱���:
    				<select name="isupchebeasong">
    				<option value="">��ü
    				<option value="Y" <%= chkIIF(isupchebeasong="Y","selected","") %> >��ü���
    				<option value="N" <%= chkIIF(isupchebeasong="N","selected","") %> >�ٹ�
    				</select>
    				&nbsp;&nbsp;

    				����:
    				<select name="divcd">
    				<option value="">��ü
    				<option value="A008" <%= chkIIF(divcd="A008","selected","") %> >�ֹ����(ǰ��,�������)
    				<option value="A004" <%= chkIIF(divcd="A004","selected","") %> >��ǰ(��ü���)
    				<option value="T012" <%= chkIIF(divcd="T012","selected","") %> >����/����/�±�ȯ
    				<option value="A000" <%= chkIIF(divcd="A000","selected","") %> >�±�ȯ
    				<option value="A001" <%= chkIIF(divcd="A001","selected","") %> >������߼�
    				<option value="A002" <%= chkIIF(divcd="A002","selected","") %> >���񽺹߼�

    				</select>
    				<input type="checkbox" name="casegroup" <%= chkIIF(casegroup="on","checked","") %>>�հ�� ǥ��

    		    </td>
    		</tr>
    		</table>
	    </td>
	</tr>
	</form>
</table>
<!-- ǥ ��ܹ� ��-->
<p>
	* �Ʒ� ������ ����� <font color=red>���</font> ����Ÿ�̹Ƿ� ��������Ÿ���� �Ǽ��� �����ϴ�.<br>
	* �� �Ǽ� : <font color=red><%= obrandCs.FTotalCount %> ��</font><br>
</p>

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">
<tr align="center" bgcolor="#E6E6E6">
	<td width="60">��/��</td>
	<td>�귣��ID</td>
	<td width="80">�׷��ڵ�</td>
  	<td width="60">���</td>
  	<td width="180">����</td>
  	<% if (casegroup="on") then %>
  	<td width="40">ǰ��</td>
  	<td width="40">��ǰ<br>�ҷ�</td>
  	<td width="40">���߼�</td>
  	<td width="40">��ǰ<br>�ļ�</td>
  	<td width="40">��ǰ<br>����</td>
  	<td width="40">���<br>����</td>
  	<td width="40">��<br>����</td>
  	<td width="40">��Ÿ</td>
  	<td width="40">�հ�</td>
  	<% else %>
  	<td width="60">����1</td>
  	<td width="60">����2</td>
  	<td>�Ǽ�</td>
  	<% end if %>
	<td width="100">���</td>
</tr>
<% for i=0 to obrandCs.FREsultCount-1 %>
<% nYYYYMMDD = Left(DateAdd("m",1,obrandCs.FItemList(i).FYYYYMM + "-01"),7) + "-01" %>
<tr bgcolor="#FFFFFF">
    <td align="center"><%= obrandCs.FItemList(i).FYYYYMM %></td>
    <td><a href="javascript:PopBrandInfoEdit('<%= obrandCs.FItemList(i).FMakerid %>');"><%= obrandCs.FItemList(i).FMakerid %></a></td>
    <td><a href="javascript:PopUpcheInfoEdit('')"></a></td>
    <td align="center"><%= ChkIIF(obrandCs.FItemList(i).Fisupchebeasong="Y","��ü","<font color='#AAAAAA'>�ٹ�</font>") %></td>
    <td align="left"><a href="javascript:popBrandCSAsDetailList('<%= obrandCs.FItemList(i).FYYYYMM %>-01','<%= nYYYYMMDD %>','<%= obrandCs.FItemList(i).FMakerid %>','<%= obrandCs.FItemList(i).Fdivcd %>','');"><%= obrandCs.FItemList(i).Fdivname %></a></td>
    <% if (casegroup="on") then %>
  	<td align="center"><a href="javascript:popBrandCSAsDetailList('<%= obrandCs.FItemList(i).FYYYYMM %>-01','<%= nYYYYMMDD %>','<%= obrandCs.FItemList(i).FMakerid %>','<%= obrandCs.FItemList(i).Fdivcd %>','CD05,CF05');"><%= obrandCs.FItemList(i).FCNT_1 %></a></td>
  	<td align="center"><a href="javascript:popBrandCSAsDetailList('<%= obrandCs.FItemList(i).FYYYYMM %>-01','<%= nYYYYMMDD %>','<%= obrandCs.FItemList(i).FMakerid %>','<%= obrandCs.FItemList(i).Fdivcd %>','CE01');"><%= obrandCs.FItemList(i).FCNT_2 %></a></td>
  	<td align="center"><a href="javascript:popBrandCSAsDetailList('<%= obrandCs.FItemList(i).FYYYYMM %>-01','<%= nYYYYMMDD %>','<%= obrandCs.FItemList(i).FMakerid %>','<%= obrandCs.FItemList(i).Fdivcd %>','CF01');"><%= obrandCs.FItemList(i).FCNT_3 %></a></td>
  	<td align="center"><a href="javascript:popBrandCSAsDetailList('<%= obrandCs.FItemList(i).FYYYYMM %>-01','<%= nYYYYMMDD %>','<%= obrandCs.FItemList(i).FMakerid %>','<%= obrandCs.FItemList(i).Fdivcd %>','CF02');"><%= obrandCs.FItemList(i).FCNT_4 %></a></td>
  	<td align="center"><a href="javascript:popBrandCSAsDetailList('<%= obrandCs.FItemList(i).FYYYYMM %>-01','<%= nYYYYMMDD %>','<%= obrandCs.FItemList(i).FMakerid %>','<%= obrandCs.FItemList(i).Fdivcd %>','CF03,CF04');"><%= obrandCs.FItemList(i).FCNT_5 %></a></td>
  	<td align="center"><a href="javascript:popBrandCSAsDetailList('<%= obrandCs.FItemList(i).FYYYYMM %>-01','<%= nYYYYMMDD %>','<%= obrandCs.FItemList(i).FMakerid %>','<%= obrandCs.FItemList(i).Fdivcd %>','CF06,CG01');"><%= obrandCs.FItemList(i).FCNT_6 %></a></td>
  	<td align="center"><a href="javascript:popBrandCSAsDetailList('<%= obrandCs.FItemList(i).FYYYYMM %>-01','<%= nYYYYMMDD %>','<%= obrandCs.FItemList(i).FMakerid %>','<%= obrandCs.FItemList(i).Fdivcd %>','CD01,CB04');"><%= obrandCs.FItemList(i).FCNT_7 %></a></td>
  	<td align="center"><a href="javascript:popBrandCSAsDetailList('<%= obrandCs.FItemList(i).FYYYYMM %>-01','<%= nYYYYMMDD %>','<%= obrandCs.FItemList(i).FMakerid %>','<%= obrandCs.FItemList(i).Fdivcd %>','CE02,CE04,CE03,CG02,CG03');"><%= (obrandCs.FItemList(i).Fcnt - (obrandCs.FItemList(i).FCNT_1 + obrandCs.FItemList(i).FCNT_2 + obrandCs.FItemList(i).FCNT_3 + obrandCs.FItemList(i).FCNT_4 + obrandCs.FItemList(i).FCNT_5 + obrandCs.FItemList(i).FCNT_6 + obrandCs.FItemList(i).FCNT_7)) %></a></td>
  	<td align="center"><a href="javascript:popBrandCSAsDetailList('<%= obrandCs.FItemList(i).FYYYYMM %>-01','<%= nYYYYMMDD %>','<%= obrandCs.FItemList(i).FMakerid %>','<%= obrandCs.FItemList(i).Fdivcd %>','');"><%= obrandCs.FItemList(i).Fcnt %></a></td>
  	<% else %>
    <td><%= obrandCs.FItemList(i).Fgubun01name %></td>
    <td><%= obrandCs.FItemList(i).Fgubun02name %></td>
    <td align="center"><%= obrandCs.FItemList(i).Fcnt %></td>
    <% end if %>
	<td align="center">
		<% if (casegroup="on") then %>
			<input type="button" class="button" value="��" onClick="popBrandCSAsDetailList('<%= obrandCs.FItemList(i).FYYYYMM %>-01','<%= nYYYYMMDD %>','<%= obrandCs.FItemList(i).FMakerid %>','<%= obrandCs.FItemList(i).Fdivcd %>','')">
			<input type="button" class="button" value="���" onClick="popBrandCSSum('<%= obrandCs.FItemList(i).FYYYYMM %>-01','<%= nYYYYMMDD %>','<%= obrandCs.FItemList(i).FMakerid %>','<%= obrandCs.FItemList(i).Fdivcd %>','')">
		<% end if %>
	</td>
</tr>
<% next %>

<% 'if (casegroup<>"on") then %>
    <tr bgcolor="#FFFFFF" height="30">
        <td colspan="16" align="center">
        <!-- ������ ���� -->
    	<%
    		if obrandCs.HasPreScroll then
    			Response.Write "<a href='javascript:goPage(" & obrandCs.StartScrollPage-1 & ")'>[pre]</a> &nbsp;"
    		else
    			Response.Write "[pre] &nbsp;"
    		end if

    		for i=0 + obrandCs.StartScrollPage to obrandCs.FScrollCount + obrandCs.StartScrollPage - 1

    			if i>obrandCs.FTotalpage then Exit for

    			if CStr(page)=CStr(i) then
    				Response.Write " <font color='red'>[" & i & "]</font> "
    			else
    				Response.Write " <a href='javascript:goPage(" & i & ")'>[" & i & "]</a> "
    			end if

    		next

    		if obrandCs.HasNextScroll then
    			Response.Write "&nbsp; <a href='javascript:goPage(" & i & ")'>[next]</a>"
    		else
    			Response.Write "&nbsp; [next]"
    		end if
    	%>
    	<!-- ������ �� -->
        </td>
    </tr>
<% 'end if %>
</table>

<script language='javascript'>
function chkComp(bool){
    document.frm.yyyy1.disabled = !(bool);
    document.frm.mm1.disabled = !(bool);

}

function getOnload(){
    chkComp(<%=ChkIIF(ck_date="on","true","false") %>);
}

window.onload = getOnload;
</script>
<%
set obrandCs = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
