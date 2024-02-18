<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/outmall_function.asp"-->
<!-- #include virtual="/lib/classes/extjungsan/extjungsancls.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<%
dim research, page
dim sellsite, searchfield, searchtext, diffType
'dim yyyy, mm, yyyymm, yyyymm_prev, yyyymm_next
'dim yyyy1,mm1

Dim i

research = requestCheckvar(request("research"),10)
page 	 = requestCheckvar(request("page"),10)

' yyyy1   = requestCheckvar(request("yyyy1"),4)
' mm1     = requestCheckvar(request("mm1"),2)

sellsite		= request("sellsite")
searchfield 	= request("searchfield")
searchtext 		= Replace(Replace(request("searchtext"), "'", ""), Chr(34), "")
diffType 		= request("diffType")

if (page="") then page = 1
if (diffType="") then diffType = "S"


' if (yyyy1="") then
' 	yyyy1 = Cstr(Year(now()))
' 	mm1 = Cstr(Month(now()) - 2)
' end if

' yyyymm = yyyy1 + "-" & mm1
' yyyymm_prev = Left(DateSerial(yyyy1,(mm1 - 1), 1), 7)
' yyyymm_next = Left(DateSerial(yyyy1,(mm1 + 1), 1), 7)


Dim oCExtJungsan
set oCExtJungsan = new CExtJungsan
	oCExtJungsan.FPageSize = 100
	oCExtJungsan.FCurrPage = page
	oCExtJungsan.FRectSellSite = sellsite
	oCExtJungsan.FRectDiffType = diffType

	'oCExtJungsan.FRectYYYYMM = yyyymm
	'oCExtJungsan.FRectSearchField = searchfield
	'oCExtJungsan.FRectSearchText = searchtext

    oCExtJungsan.GetExtJungsanDiff


%>

<script language='javascript'>

function NextPage(page){
    document.frm.page.value = page;
    document.frm.submit();
}

function popExtSiteJungsanData() {
    var window_width = 500;
    var window_height = 250;

    var popwin = window.open("/admin/maechul/extjungsandata/popRegExtJungsanDataFile.asp","popExtSiteJungsanData","width=" + window_width + " height=" + window_height + " left=0 top=0 scrollbars=yes resizable=yes status=yes");

	popwin.focus();
}



function popMiMapExtjungsan(sellsite,yyyy1,mm1,dd1,yyyy2,mm2,dd2,jungsantype){
	var iurl = "/admin/maechul/extjungsandata/extJungsanDataList.asp";
	iurl += "?menupos=1652&page=1&sellsite="+sellsite;
	iurl += "&yyyy1="+yyyy1+"&mm1="+mm1+"&dd1="+dd1;
	iurl += "&yyyy2="+yyyy2+"&mm2="+mm2+"&dd2="+dd2;
	iurl += "&mimap=on&jungsantype="+jungsantype;

	var popwin = window.open(iurl,"popextJungsanDataList","width=1200 height=850 scrollbars=yes resizable=yes status=yes");

	popwin.focus();

}

function popFixedErrExtjungsanDTL(sellsite,yyyy1,mm1,jungsantype,ierrtp,iaccerrtype){
	var iurl = "/admin/maechul/extjungsandata/extJungsanFixedErrDetail.asp";
	iurl += "?menupos=1652&page=1&sellsite="+sellsite;
	iurl += "&yyyy1="+yyyy1+"&mm1="+mm1;
	iurl += "&jungsantype="+jungsantype;
	iurl += "&errtp="+ierrtp;
	iurl += "&accerrtype="+iaccerrtype;

	var popwin = window.open(iurl,"popextJungsanFixedErrDetail","width=1200 height=850 scrollbars=yes resizable=yes status=yes");

	popwin.focus();

}

function popErrExtjungsanDTL(sellsite,yyyy1,mm1,dd1,yyyy2,mm2,dd2,jungsantype,ierrtp,ionlyErrNoExists){
	var iurl = "/admin/maechul/extjungsandata/extJungsanErrDetail.asp";
	iurl += "?menupos=1652&page=1&sellsite="+sellsite;
	iurl += "&yyyy1="+yyyy1+"&mm1="+mm1+"&dd1="+dd1;
	iurl += "&yyyy2="+yyyy2+"&mm2="+mm2+"&dd2="+dd2;
	iurl += "&jungsantype="+jungsantype;
	iurl += "&errtp="+ierrtp;
	iurl += "&onlyErrNoExists="+ionlyErrNoExists;

	var popwin = window.open(iurl,"popextJungsanErrDetail","width=1200 height=850 scrollbars=yes resizable=yes status=yes");

	popwin.focus();

}

function jsExtJungsanDiffMake(sellsite,yyyymm) {
	var frm = document.frmAct;

	if (confirm(sellsite + " "+yyyymm+" (��)�ۼ��Ͻðڽ��ϱ�?") == true) {
		frm.mode.value = "extjungsandiffmake";
		frm.sellsite.value = sellsite;
		frm.yyyymm.value = yyyymm;

		frm.submit();
	}
}


function jsExtJungsanErrMake(sellsite,yyyymm) {
	var frm = document.frmAct;

	if (confirm(sellsite + " "+yyyymm+" Err �ۼ��Ͻðڽ��ϱ�?") == true) {
		frm.mode.value = "extjungsanerrmake";
		frm.sellsite.value = sellsite;
		frm.yyyymm.value = yyyymm;

		frm.submit();
	}
}

function jsExtJungsanDiffMakeDetail(sellsite){
	var frm = document.frmAct;

	if (confirm(sellsite + " ���� ���� �ۼ��Ͻðڽ��ϱ�?") == true) {
		frm.mode.value = "extjungsanaccDetailmake";
		frm.sellsite.value = sellsite;
		frm.yyyymm.value = "";
		frm.submit();
	}
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
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">�˻�<br>����</td>
	<td align="left">
		&nbsp;
		���޸�:	<%' = getJungsanXsiteComboHTML("sellsite",sellsite,"") %>
				<% call drawOutmallSelectBox("sellsite",sellsite) %>
		<% if (FALSE) then %>
		&nbsp;
		�����:
		<% DrawYMBox yyyy1,mm1 %>
		<% end if %>
		&nbsp;
		��ȸ����:

		<input type="radio" name="diffType" value="S" <% if (diffType = "S") then %>checked<% end if %> > �����������
		<input type="radio" name="diffType" value="T" <% if (diffType = "T") then %>checked<% end if %> > TEN�������

		<% if (FALSE) then %>
		<input type="radio" name="diffType" value="DIF" <% if (diffType = "DIF") then %>checked<% end if %> > ��������
		<input type="radio" name="diffType" value="TOT" <% if (diffType = "TOT") then %>checked<% end if %> > ��ü����
		<input type="radio" name="diffType" value="SUM" <% if (diffType = "SUM") then %>checked<% end if %> > �հ賻��
		<% end if %>
	</td>
	<td rowspan="4" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="�˻�" onClick="javascript:document.frm.submit();">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" height="30">
	<td align="left">
		&nbsp;
		<% if (FALSE) then %>
		* �˻����� :
		<select class="select" name="searchfield">
			<option value=""></option>
			<option value="OrgOrderserial" <% if (searchfield = "OrgOrderserial") then %>selected<% end if %> >���ֹ���ȣ</option>
		</select>
		<input type="text" class="text" name="searchtext" size="30" value="<%= searchtext %>">
		<% end if %>
	</td>
</tr>
</form>
</table>
<!-- �˻� �� -->

<%

if (sellsite = "") then
	Response.write "<h5>���޸��� �����ϼ���</h5>"
end if

%>

<p>

<!-- �׼� ���� -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
	<td align="left">

	</td>
	<td align="right">

	</td>
</tr>
</table>
<!-- �׼� �� -->

<p>

<!-- ����Ʈ ���� -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr height="25" bgcolor="FFFFFF">
	<td colspan="<%=CHKIIF(diffType="S","21","20")%>">
		�˻���� : <b><%= oCExtJungsan.FTotalcount %></b>
		&nbsp;
		������ : <b><%= page %> / <%= oCExtJungsan.FTotalPage %></b>
	</td>
</tr>
<% if diffType="S" then %>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="70">���</td>
	<td width="100">�ڻ������<br>(�������)</td>
	<td width="100">�ڻ����(��ǰ)</td>
	<td width="90">�ڻ����(��ۺ�)</td>
	<td width="1"></td>
	<td width="100">���޸�����<br>(���������Է±���)</td>
	<td width="100">���޸���(��ǰ)</td>
	<td width="90">���޸���(��ۺ�)</td>
	<td width="1"></td>
	<td width="100">�������(��ǰ)<br>(�ڻ�-����)</td>
	<td width="90">�������(��ۺ�)<br>(�ڻ�-����)</td>
	<td width="1"></td>
	<td width="100">��������(��ǰ)<br>(�ڻ�-����)</td>
	<td width="100">��������(��ۺ�)<br>(�ڻ�-����)</td>
	<td width="120">����������Ʈ</td>
	<td width="80">���������<br>(�ڻ�-����)</td>
	<td width="80">�̹ݿ�����<br>(�ڻ�-����)</td>
	<td width="80">���Fix����<br>(�ڻ�-����)</td>
	<td width="80">���Fix����<br>(�ڻ�-����)��ǰ</td>
	<td width="90">���Fix����<br>(�ڻ�-����)��ǰ<br>�����ʿ�</td>
	<td>
		���
		<% if oCExtJungsan.FresultCount>0 then %>
			<% if oCExtJungsan.FItemList(0).Fyyyymm<LEFT(now(),7) then %>
			<% if (sellsite<>"") then %>
			<br><input type="button" value="<%=LEFT(now(),7)%>�ۼ�" onClick="jsExtJungsanDiffMake('<%=sellsite%>','<%=LEFT(now(),7)%>')">
			<% end if %>
			<% end if %>
		<% end if %>
	</td>
</tr>
<% for i=0 to oCExtJungsan.FresultCount -1 %>
<tr align="right" bgcolor="FFFFFF" onmouseover=this.style.background="F1F1F1"; onmouseout=this.style.background="FFFFFF";>
	<td align="center" ><%= oCExtJungsan.FItemList(i).Fyyyymm %></td>
	<td><%= FormatNumber(oCExtJungsan.FItemList(i).FTMeachulItem+oCExtJungsan.FItemList(i).FTMeachulDLV,0) %></td>
	<td><%= FormatNumber(oCExtJungsan.FItemList(i).FTMeachulItem,0) %></td>
	<td><%= FormatNumber(oCExtJungsan.FItemList(i).FTMeachulDLV,0) %></td>
	<td></td>
	<td><%= FormatNumber(oCExtJungsan.FItemList(i).FXMeachulItem+oCExtJungsan.FItemList(i).FXMeachulDLV,0) %></td>
	<td><%= FormatNumber(oCExtJungsan.FItemList(i).FXMeachulItem,0) %></td>
	<td><%= FormatNumber(oCExtJungsan.FItemList(i).FXMeachulDLV,0) %></td>
	<td></td>
	<td><%= FormatNumber(oCExtJungsan.FItemList(i).FmonthItemDiff,0) %></td>
	<td><%= FormatNumber(oCExtJungsan.FItemList(i).FmonthdlvDiff,0) %></td>
	<td></td>
	<td><%= FormatNumber(oCExtJungsan.FItemList(i).FdiffITEMsum,0) %></td>
	<td><%= FormatNumber(oCExtJungsan.FItemList(i).FdiffDlvsum,0) %></td>

	<td align="center" ><%= LEFT(oCExtJungsan.FItemList(i).FupdDt,16) %></td>
	<td>
		<% if NOT isNULL(oCExtJungsan.FItemList(i).FMonthDiffSum) then %>
		<a href="#" onClick="popFixedErrExtjungsanDTL('<%=sellsite%>','<%= LEFT(oCExtJungsan.FItemList(i).Fyyyymm,4) %>','<%= RIGHT(oCExtJungsan.FItemList(i).Fyyyymm,2) %>','','','');return false;"><%= FormatNumber(oCExtJungsan.FItemList(i).FMonthDiffSum,0) %></a>
			<% if oCExtJungsan.FItemList(i).getSumVsDtlDiffSum<>0 then %>
				<br><font color=gray><%=FormatNumber(oCExtJungsan.FItemList(i).getSumVsDtlDiffSum,0) %></font>
			<% end if %>
		<% end if %>
	</td>
	<td>
		<a href="#" onClick="popFixedErrExtjungsanDTL('<%=sellsite%>','<%= LEFT(oCExtJungsan.FItemList(i).Fyyyymm,4) %>','<%= RIGHT(oCExtJungsan.FItemList(i).Fyyyymm,2) %>','','1','');return false;"><%= FormatNumber(oCExtJungsan.FItemList(i).FMonthnotAssignErr,0) %></a>
	</td>
	<td>
		<a href="#" onClick="popFixedErrExtjungsanDTL('<%=sellsite%>','<%= LEFT(oCExtJungsan.FItemList(i).Fyyyymm,4) %>','<%= RIGHT(oCExtJungsan.FItemList(i).Fyyyymm,2) %>','','3','');return false;"><%= FormatNumber(oCExtJungsan.FItemList(i).FMonthErrAsignSum,0) %></a>
	</td>
	<td>
		<a href="#" onClick="popFixedErrExtjungsanDTL('<%=sellsite%>','<%= LEFT(oCExtJungsan.FItemList(i).Fyyyymm,4) %>','<%= RIGHT(oCExtJungsan.FItemList(i).Fyyyymm,2) %>','C','3','');return false;"><%= FormatNumber(oCExtJungsan.FItemList(i).FMonthErrAsignItemSum,0) %></a>
	</td>
	<td>
		<a href="#" onClick="popFixedErrExtjungsanDTL('<%=sellsite%>','<%= LEFT(oCExtJungsan.FItemList(i).Fyyyymm,4) %>','<%= RIGHT(oCExtJungsan.FItemList(i).Fyyyymm,2) %>','C','3','3');return false;"><%= FormatNumber(oCExtJungsan.FItemList(i).FMonthErrAsignItemSumReqCheck,0) %></a>
	</td>
	<td align="center" >
	<% if (sellsite<>"") then %>
	<% if (LEFT(dateadd("m",-3,now()),7)<=oCExtJungsan.FItemList(i).Fyyyymm) then %>
	<input type="button" value="���ۼ�" onClick="jsExtJungsanDiffMake('<%=sellsite%>','<%=oCExtJungsan.FItemList(i).Fyyyymm%>')">

		<% if (oCExtJungsan.FItemList(i).Fyyyymm>="2020-01") then %>
		<input type="button" value="Fix����" onClick="jsExtJungsanErrMake('<%=sellsite%>','<%=oCExtJungsan.FItemList(i).Fyyyymm%>')">
		<% end if %>
	<% end if %>
	<% end if %>

	</td>
</tr>
<% next %>
<% else %>
<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
	<td width="70">���</td>
	<td width="120">�ڻ������<br>(�������)</td>
	<td width="110">�ڻ����(��ǰ)</td>
	<td width="100">�ڻ����(��ۺ�)</td>
	<td width="1"></td>
	<td width="120">���޸�����<br>(�̸��γ���)</td>
	<td width="110">���޸���(��ǰ)<br>(�̸��γ���)</td>
	<td width="110">���޸���(��ۺ�)<br>(�̸��γ���)</td>
	<td width="1"></td>
	<td width="120">������ο���(��ǰ)<br>(�ڻ�-����)</td>
	<td width="120">������ο���(��ۺ�)<br>(�ڻ�-����)</td>

	<td width="110"><font color="#AAAAAA">������ο���<br>(��ǰ ������)</font></td>
	<td width="100"><font color="#AAAAAA">������ο���<br>(��ۺ� ������)</font></td>

	<td width="110"><font color="#AAAAAA">������ο���<br>(��ǰ ��������)</font></td>
	<td width="100"><font color="#AAAAAA">������ο���<br>(��ۺ� ��������)</font></td>

	<td width="1"></td>
	<td width="120">�������ο���(��ǰ)<br>(�ڻ�-����-�̸���)</td>
	<td width="120">�������ο���(��ۺ�)<br>(�ڻ�-����-�̸���)</td>
	<td width="200">����������Ʈ</td>
	<td>
		���
		<% if (sellsite<>"") then %>
		<% if oCExtJungsan.FresultCount>0 then %>
			<br><input type="button" value="���ۼ�" onClick="jsExtJungsanDiffMakeDetail('<%=sellsite%>')">
		<% end if %>
		<% end if %>
	</td>
</tr>
<% for i=0 to oCExtJungsan.FresultCount -1 %>
<tr align="right" bgcolor="FFFFFF" onmouseover=this.style.background="F1F1F1"; onmouseout=this.style.background="FFFFFF";>
	<td align="center" ><%= oCExtJungsan.FItemList(i).Fyyyymm %></td>
	<td><%= FormatNumber(oCExtJungsan.FItemList(i).FTMeachulItem+oCExtJungsan.FItemList(i).FTMeachulDLV,0) %></td>
	<td><%= FormatNumber(oCExtJungsan.FItemList(i).FTMeachulItem,0) %></td>
	<td><%= FormatNumber(oCExtJungsan.FItemList(i).FTMeachulDLV,0) %></td>
	<td></td>
	<td><a href="#" onClick="popMiMapExtjungsan('<%=sellsite%>','<%= LEFT(oCExtJungsan.FItemList(i).Fyyyymm,4) %>','<%= RIGHT(oCExtJungsan.FItemList(i).Fyyyymm,2) %>','01','<%= LEFT(oCExtJungsan.FItemList(i).Fyyyymm,4) %>','<%= RIGHT(oCExtJungsan.FItemList(i).Fyyyymm,2) %>','<%= RIGHT(dateadd("d",-1,dateadd("m",1,oCExtJungsan.FItemList(i).Fyyyymm+"-01")),2) %>','');return false;"><%= FormatNumber(oCExtJungsan.FItemList(i).FXMeachulItem+oCExtJungsan.FItemList(i).FXMeachulDLV,0) %></a></td>
	<td><a href="#" onClick="popMiMapExtjungsan('<%=sellsite%>','<%= LEFT(oCExtJungsan.FItemList(i).Fyyyymm,4) %>','<%= RIGHT(oCExtJungsan.FItemList(i).Fyyyymm,2) %>','01','<%= LEFT(oCExtJungsan.FItemList(i).Fyyyymm,4) %>','<%= RIGHT(oCExtJungsan.FItemList(i).Fyyyymm,2) %>','<%= RIGHT(dateadd("d",-1,dateadd("m",1,oCExtJungsan.FItemList(i).Fyyyymm+"-01")),2) %>','C');return false;"><%= FormatNumber(oCExtJungsan.FItemList(i).FXMeachulItem,0) %></a></td>
	<td><a href="#" onClick="popMiMapExtjungsan('<%=sellsite%>','<%= LEFT(oCExtJungsan.FItemList(i).Fyyyymm,4) %>','<%= RIGHT(oCExtJungsan.FItemList(i).Fyyyymm,2) %>','01','<%= LEFT(oCExtJungsan.FItemList(i).Fyyyymm,4) %>','<%= RIGHT(oCExtJungsan.FItemList(i).Fyyyymm,2) %>','<%= RIGHT(dateadd("d",-1,dateadd("m",1,oCExtJungsan.FItemList(i).Fyyyymm+"-01")),2) %>','D');return false;"><%= FormatNumber(oCExtJungsan.FItemList(i).FXMeachulDLV,0) %></a></td>
	<td></td>
	<td><a href="#" onClick="popErrExtjungsanDTL('<%=sellsite%>','<%= LEFT(oCExtJungsan.FItemList(i).Fyyyymm,4) %>','<%= RIGHT(oCExtJungsan.FItemList(i).Fyyyymm,2) %>','01','<%= LEFT(oCExtJungsan.FItemList(i).Fyyyymm,4) %>','<%= RIGHT(oCExtJungsan.FItemList(i).Fyyyymm,2) %>','<%= RIGHT(dateadd("d",-1,dateadd("m",1,oCExtJungsan.FItemList(i).Fyyyymm+"-01")),2) %>','C','','');return false;"><%= FormatNumber(oCExtJungsan.FItemList(i).FmonthItemDiff,0) %></a></td>
	<td><a href="#" onClick="popErrExtjungsanDTL('<%=sellsite%>','<%= LEFT(oCExtJungsan.FItemList(i).Fyyyymm,4) %>','<%= RIGHT(oCExtJungsan.FItemList(i).Fyyyymm,2) %>','01','<%= LEFT(oCExtJungsan.FItemList(i).Fyyyymm,4) %>','<%= RIGHT(oCExtJungsan.FItemList(i).Fyyyymm,2) %>','<%= RIGHT(dateadd("d",-1,dateadd("m",1,oCExtJungsan.FItemList(i).Fyyyymm+"-01")),2) %>','D','','');return false;"><%= FormatNumber(oCExtJungsan.FItemList(i).FmonthdlvDiff,0) %></a></td>

	<td><a href="#" onClick="popErrExtjungsanDTL('<%=sellsite%>','<%= LEFT(oCExtJungsan.FItemList(i).Fyyyymm,4) %>','<%= RIGHT(oCExtJungsan.FItemList(i).Fyyyymm,2) %>','01','<%= LEFT(oCExtJungsan.FItemList(i).Fyyyymm,4) %>','<%= RIGHT(oCExtJungsan.FItemList(i).Fyyyymm,2) %>','<%= RIGHT(dateadd("d",-1,dateadd("m",1,oCExtJungsan.FItemList(i).Fyyyymm+"-01")),2) %>','C','1','');return false;"><font color="#AAAAAA"><%= FormatNumber(oCExtJungsan.FItemList(i).FmonthItemDiffMapErr,0) %></font></a></td>
	<td><a href="#" onClick="popErrExtjungsanDTL('<%=sellsite%>','<%= LEFT(oCExtJungsan.FItemList(i).Fyyyymm,4) %>','<%= RIGHT(oCExtJungsan.FItemList(i).Fyyyymm,2) %>','01','<%= LEFT(oCExtJungsan.FItemList(i).Fyyyymm,4) %>','<%= RIGHT(oCExtJungsan.FItemList(i).Fyyyymm,2) %>','<%= RIGHT(dateadd("d",-1,dateadd("m",1,oCExtJungsan.FItemList(i).Fyyyymm+"-01")),2) %>','D','1','');return false;"><font color="#AAAAAA"><%= FormatNumber(oCExtJungsan.FItemList(i).FmonthdlvDiffTMapErr,0) %></font></a></td>

	<td><a href="#" onClick="popErrExtjungsanDTL('<%=sellsite%>','<%= LEFT(oCExtJungsan.FItemList(i).Fyyyymm,4) %>','<%= RIGHT(oCExtJungsan.FItemList(i).Fyyyymm,2) %>','01','<%= LEFT(oCExtJungsan.FItemList(i).Fyyyymm,4) %>','<%= RIGHT(oCExtJungsan.FItemList(i).Fyyyymm,2) %>','<%= RIGHT(dateadd("d",-1,dateadd("m",1,oCExtJungsan.FItemList(i).Fyyyymm+"-01")),2) %>','C','','on');return false;"><font color="#AAAAAA"><%= FormatNumber(oCExtJungsan.FItemList(i).FmonthItemDiffNoExists,0) %></font></a></td>
	<td><a href="#" onClick="popErrExtjungsanDTL('<%=sellsite%>','<%= LEFT(oCExtJungsan.FItemList(i).Fyyyymm,4) %>','<%= RIGHT(oCExtJungsan.FItemList(i).Fyyyymm,2) %>','01','<%= LEFT(oCExtJungsan.FItemList(i).Fyyyymm,4) %>','<%= RIGHT(oCExtJungsan.FItemList(i).Fyyyymm,2) %>','<%= RIGHT(dateadd("d",-1,dateadd("m",1,oCExtJungsan.FItemList(i).Fyyyymm+"-01")),2) %>','D','','on');return false;"><font color="#AAAAAA"><%= FormatNumber(oCExtJungsan.FItemList(i).FmonthdlvDiffTNoExists,0) %></font></a></td>

	<td></td>
	<td><%= FormatNumber(oCExtJungsan.FItemList(i).FdiffITEMsum,0) %></td>
	<td><%= FormatNumber(oCExtJungsan.FItemList(i).FdiffDlvsum,0) %></td>

	<td align="center" ><%= oCExtJungsan.FItemList(i).FupdDt %></td>
	<td align="center" ></td>
</tr>
<% next %>
<% end if %>

<tr height="25" bgcolor="FFFFFF">
	<td colspan="<%=CHKIIF(diffType="S","21","20")%>" align="center">
	<% if (FALSE) then %>
		<% if oCExtJungsan.HasPreScroll then %>
		<a href="javascript:NextPage('<%= oCExtJungsan.StartScrollPage-1 %>')">[pre]</a>
		<% else %>
			[pre]
		<% end if %>

		<% for i=0 + oCExtJungsan.StartScrollPage to oCExtJungsan.FScrollCount + oCExtJungsan.StartScrollPage - 1 %>
			<% if i>oCExtJungsan.FTotalpage then Exit for %>
			<% if CStr(page)=CStr(i) then %>
			<font color="red">[<%= i %>]</font>
			<% else %>
			<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
			<% end if %>
		<% next %>

		<% if oCExtJungsan.HasNextScroll then %>
			<a href="javascript:NextPage('<%= i %>')">[next]</a>
		<% else %>
			[next]
		<% end if %>
	<% end if %>
	</td>
</tr>
</table>

<form name="frmAct" method="post" action="extJungsan_process.asp">
<input type="hidden" name="mode" value="">
<input type="hidden" name="sellsite" value="">
<input type="hidden" name="yyyymm" value="">
</form>

<%
set oCExtJungsan = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
