<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/cscenter/resending_reportcls.asp"-->

<%
dim page, research

page	= req("page",1)


Dim ck_date1, ck_date2, finishDate1, finishDate2, regDate1, regDate2, gubun01, gubun02
dim isupchebeasong, divcd
dim finishuser, reguserid

research	= req("research","")
ck_date1	= req("ck_date1","")
ck_date2	= req("ck_date2","")

finishDate1	= req("finishDate1",Date())
finishDate2	= req("finishDate2",Date())
regDate1	= req("regDate1",Date())
regDate2	= req("regDate2",Date())

divCd	= req("divCd","")
gubun01	= req("gubun01","")
gubun02	= req("gubun02","")
isUpcheBeasong	= req("isUpcheBeasong","")
finishuser	= req("finishuser","")
reguserid	= req("reguserid","")

dim obj
set obj = new CReportMaster
obj.FPageSize = 50
obj.FCurrPage = page

obj.FRectFinishUser = finishuser
if (ck_date2="on") then
    obj.FRectRegStart = regDate1
    obj.FRectRegEnd = regDate2
end if
obj.FRectRegUserID = reguserid

obj.getCsListView2 finishDate1, finishDate2, divCd, gubun01, gubun02

dim i
%>
<script language='javascript'>
function goPage(page){
    frm.page.value = page;
    frm.submit();
}

function jsPopCal(fName,sName)
{
	var winCal;
	winCal = window.open('/lib/common_cal.asp?FN='+fName+'&DN='+sName,'pCal','width=250, height=200');
	winCal.focus();
}


window.onload = function()
{
	document.title = "CSó���Ϸ�����ȸ";
	getGubun02Options('<%=gubun02%>');
}

function getGubun02Options(gubun02)
{
	var f = document.frm;

	switch (f.gubun01.value)
	{
	case "C004":
		var arr =
		[
			["CD01", "�ܼ�����"],
			["CD03", "���ֹ�"],
			["CD04", "������"],
			["CD05", "ǰ��"],
			["CD99", "��Ÿ"]
		];
		break;
	case "C005":
		var arr =
		[
			["CE01", "��ǰ�ҷ�"],
			["CE02", "��ǰ�Ҹ���"],
			["CE03", "��ǰ��Ͽ���"],
			["CE04", "��ǰ����ҷ�"],
			["CE05", "�̺�Ʈ�����"],
			["CE99", "��Ÿ"]
		];
		break;
	case "C006":
		var arr =
		[
			["CF01", "���߼�"],
			["CF02", "��ǰ�ļ�"],
			["CF03", "���Ż�ǰ����"],
			["CF04", "����ǰ����"],
			["CF05", "��ǰǰ��"],
			["CF06", "�������"],
			["CF99", "��Ÿ"]
		];
		break;
	case "C007":
		var arr =
		[
			["CG01", "�������"],
			["CG02", "�ù���ļ�"],
			["CG03", "�ù��н�"]
		];
		break;
	default:
		var arr = [];
		break;
	}

	f.gubun02.length = 1;
	for (i=0;i<arr.length ;i++ )
	{
		var newOpt = document.createElement("OPTION");
		newOpt.value = arr[i][0];
		newOpt.text  = arr[i][1];
		f.gubun02.options.add(newOpt);
	}

	if (gubun02)
		f.gubun02.value = gubun02;

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
    				<input type="checkbox" name="ck_date1" checked disabled >
					ó���Ϸ��� : <input type="text" size="10" name="finishDate1" value="<%= finishDate1 %>" onClick="jsPopCal('frm','finishDate1');" class="text" style="cursor:hand;">
					~<input type="text" size="10" name="finishDate2" value="<%= finishDate2 %>" onClick="jsPopCal('frm','finishDate2');" class="text" style="cursor:hand;">
    				&nbsp;&nbsp;
					ó���� : <input type="text" class="text" size="15" name="finishuser" value="<%= finishuser %>">
                    &nbsp;&nbsp;
                    <input type="checkbox" name="ck_date2" <%= CHKIIF(ck_date2="on", "checked", "") %>>
					������ : <input type="text" size="10" name="regDate1" value="<%= regDate1 %>" onClick="jsPopCal('frm','regDate1');" class="text" style="cursor:hand;">
					~<input type="text" size="10" name="regDate2" value="<%= regDate2 %>" onClick="jsPopCal('frm','regDate2');" class="text" style="cursor:hand;">
    				&nbsp;&nbsp;
					������ : <input type="text" class="text" size="15" name="reguserid" value="<%= reguserid %>">
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

    				����:
    				<select name="divCd">
    				<option value="">��ü
						<option value="A000" <% if (divcd = "A000") then response.write "selected" end if %>>�±�ȯ���</option>
						<option value="A001" <% if (divcd = "A001") then response.write "selected" end if %>>������߼�</option>
						<option value="A002" <% if (divcd = "A002") then response.write "selected" end if %>>���񽺹߼�</option>
						<option value="A003" <% if (divcd = "A003") then response.write "selected" end if %>>ȯ�ҿ�û</option>
						<option value="A004" <% if (divcd = "A004") then response.write "selected" end if %>>��ǰ����(��ü���)</option>
						<option value="A005" <% if (divcd = "A005") then response.write "selected" end if %>>�ܺθ�ȯ�ҿ�û</option>
						<option value="A006" <% if (divcd = "A006") then response.write "selected" end if %>>�������ǻ���</option>
						<option value="A007" <% if (divcd = "A007") then response.write "selected" end if %>>�ſ�ī��/��ü��ҿ�û</option>
						<option value="A008" <% if (divcd = "A008") then response.write "selected" end if %>>�ֹ����</option>
						<option value="A009" <% if (divcd = "A009") then response.write "selected" end if %>>��Ÿ����(�޸�)</option>
						<option value="A010" <% if (divcd = "A010") then response.write "selected" end if %>>ȸ����û(�ٹ����ٹ��)</option>
						<option value="A011" <% if (divcd = "A011") then response.write "selected" end if %>>�±�ȯȸ��(�ٹ����ٹ��)</option>
						<option value="A700" <% if (divcd = "A700") then response.write "selected" end if %>>��ü��Ÿ����</option>
						<option value="A900" <% if (divcd = "A900") then response.write "selected" end if %>>�ֹ���������</option>
    				</select>
    				&nbsp;&nbsp;

    				�������� :
    				<select name="gubun01" onchange="getGubun02Options();">
	    				<option value="">��ü</option>
	    				<option value="C004" <%=chkIIF(gubun01="C004","selected","")%>>����</option>
	    				<option value="C005" <%=chkIIF(gubun01="C005","selected","")%>>��ǰ����</option>
	    				<option value="C006" <%=chkIIF(gubun01="C006","selected","")%>>��������</option>
	    				<option value="C007" <%=chkIIF(gubun01="C007","selected","")%>>�ù�����</option>
    				</select>
    				<select name="gubun02">
	    				<option value="">��</option>
    				</select>


    		    </td>
    		</tr>
    		</table>
	    </td>
	</tr>
	</form>
</table>
<!-- ǥ ��ܹ� ��-->
<p>

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">
<tr bgcolor="#FFFFFF">
    <td colspan="11" align="right">�� <%= obj.FTotalCount %> �� <%= page %>/<%= obj.FTotalPage %> &nbsp;</td>
</tr>
<tr align="center" bgcolor="#E6E6E6">
	<!-- td width="60">ID</td -->
	<td width="150">����</td>
	<td width="150">����</td>
	<td width="100">�ֹ���ȣ</td>
	<td width="">��������</td>
	<td width="80">������</td>
	<td width="80">������</td>
    <td width="80">ó����</td>
    <td width="80">ó����</td>
</tr>
<% if obj.FResultCount<1 then %>
<tr bgcolor="#FFFFFF">
    <td align="center" colspan="11">
       [�˻� ����� �����ϴ�.]
    </td>
</tr>
<% else %>
<% for i=0 To obj.FREsultCount -1 %>
<tr bgcolor="#FFFFFF">
    <!--td ><%= obj.FItemList(i).FID %></td -->
    <td align="center"><%= obj.FItemList(i).Fdivcd_Name %></td>
    <td align="center"><%= obj.FItemList(i).Fgubun01_Name %> / <%= obj.FItemList(i).Fgubun02_Name %></td>
    <td align="center"><a href="javascript:PopOrderMasterWithCallRingOrderserial('<%= obj.FItemList(i).Forderserial %>');"><%= obj.FItemList(i).Forderserial %></a></td>
    <td align="left">&nbsp; <%= obj.FItemList(i).Ftitle %></td>
    <td align="center"><%= obj.FItemList(i).Freguserid %></td>
    <td ><%= Left(obj.FItemList(i).Fregdate,10) %></td>
	<td align="center"><%= obj.FItemList(i).Ffinishuser %></td>
    <td ><%= Left(obj.FItemList(i).Ffinishdate,10) %></td>
</tr>
<% next %>
<tr bgcolor="#FFFFFF">
    <td colspan="11" align="center">
        <!-- ������ ���� -->
			<% sbDisplayPaging "page="&page, obj.FTotalCount, obj.FPageSize, 10%>
    	<!-- ������ �� -->
    </td>
</tr>
<% end if %>
</table>


<%
set obj = Nothing
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
