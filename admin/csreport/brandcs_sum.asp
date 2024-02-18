<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/datamart/brandcssummaryclass.asp"-->

<%


dim ck_date
dim yyyy1,mm1,dd1, yyyy2,mm2,dd2
dim makerid, isupchebeasong, divcd, gubunNm, gubun02Arr
dim startdateStr, nextdateStr
dim page

ck_date     = RequestCheckVar(request("ck_date"),9)
yyyy1       = RequestCheckVar(request("yyyy1"),4)
mm1         = RequestCheckVar(request("mm1"),2)
dd1         = RequestCheckVar(request("dd1"),2)
yyyy2       = RequestCheckVar(request("yyyy2"),4)
mm2         = RequestCheckVar(request("mm2"),2)
dd2         = RequestCheckVar(request("dd2"),2)

makerid         = RequestCheckVar(request("makerid"),32)
isupchebeasong  = RequestCheckVar(request("isupchebeasong"),9)
divcd           = RequestCheckVar(request("divcd"),4)

gubunNm         = RequestCheckVar(request("gubunNm"),32)
page            = RequestCheckVar(request("page"),9)

if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = Cstr(day(now()))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))
if (page="") then page=1
if (gubunNm<>"") then gubun02Arr="'" + Replace(gubunNm,",","','") + "'"
startdateStr = yyyy1 + "-" + mm1 + "-" + dd1
nextdateStr = Left(CStr(DateAdd("d",Cdate(yyyy2 + "-" + mm2 + "-" + dd2),1)),10)

dim obrandCs
set obrandCs = new CBrandCSSummary
obrandCs.FPageSize = 50
obrandCs.FCurrPage = page
obrandCs.FRectMakerid   = makerid
obrandCs.FRectStartDate = startdateStr
obrandCs.FRectEndDate   = nextdateStr
obrandCs.FRectIsUpchebeasong = isupchebeasong
obrandCs.FRectDivCd     = divcd
obrandCs.FRectgubun02Arr = gubun02Arr

if (makerid<>"") then
    obrandCs.getBrandCsSUMList
end if

dim i
%>
<script language='javascript'>
function goPage(page){
    frm.page.value = page;
    frm.submit();
}

function fnSearch(frm) {
    frm.ck_date.disabled = false;
    frm.submit();
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
    				<input type="checkbox" name="ck_date" checked disabled >
    				ó���Ϸ��� : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
    				&nbsp;&nbsp;
    				�귣��: <% drawSelectBoxDesignerwithName "makerid", makerid %>


				</td>
			</tr>
			</table>
        </td>
		<td  width="50" bgcolor="<%= adminColor("gray") %>" rowspan="2">
		    <input type="button" class="button_s" value="�˻�" onClick="fnSearch(document.frm)">
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
    				&nbsp;&nbsp;

    				�󼼻��� :
    				<select name="gubunNm">
    				<option value="">��ü
    				<option value="CD05,CF05" <%= chkIIF(gubunNm="CD05,CF05","selected","") %> >ǰ��
    				<option value="CF06,CG01" <%= chkIIF(gubunNm="CE01","selected","") %> >��ǰ�ҷ�
    				<option value="CE01,CE02" <%= chkIIF(gubunNm="CF01","selected","") %> >���߼�
    				<option value="CF03,CF04,CF01" <%= chkIIF(gubunNm="CF02","selected","") %> >��ǰ�ļ�
    				<option value="CE04,CE03" <%= chkIIF(gubunNm="CF03,CF04","selected","") %> >��ǰ����
    				<option value="CF02,CG02,CG03" <%= chkIIF(gubunNm="CF06,CG01","selected","") %> >�������
    				<option value="CD01,CB04" <%= chkIIF(gubunNm="CD01,CB04","selected","") %> > ������
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

* �ִ� 100���� ��ǰ�� ǥ�õ˴ϴ�.

<p>

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">
<tr bgcolor="#FFFFFF">
    <td colspan="11" align="right">�� <%= obrandCs.FTotalCount %> �� <%= page %>/<%= obrandCs.FTotalPage %> &nbsp;</td>
</tr>
<tr align="center" bgcolor="#E6E6E6">
	<!-- td width="60">ID</td -->
	<td width="100">����</td>
	<td width="80">����</td>
	<td width="50">��ǰ�ڵ�</td>
	<td width="150">��ǰ��</td>
	<td width="40">����</td>
	<td width="40">���<br>����</td>
	<td ></td>
</tr>
<% if obrandCs.FResultCount<1 then %>
<tr bgcolor="#FFFFFF">
    <td align="center" colspan="11">
        <% if (makerid="") then %>
            <font color="blue">[�귣�� ID�� �����ϼ���.]</font>
        <% else  %>
            [�˻� ����� �����ϴ�.]
        <% end if %>
    </td>
</tr>
<% else %>
<% for i=0 to obrandCs.FREsultCount -1 %>
<tr bgcolor="#FFFFFF">
    <!--td ><%= obrandCs.FItemList(i).FID %></td -->
    <td align="center"><%= obrandCs.FItemList(i).Fdivcd_Name %></td>
    <td align="center"><%= obrandCs.FItemList(i).Fgubun02_Name %></td>
    <td align="center"><%= obrandCs.FItemList(i).FItemID %></td>
    <td ><%= obrandCs.FItemList(i).FItemName %>
    </td>
    <td align="center"><%= obrandCs.FItemList(i).FconfirmItemNo %></td>
    <td align="center"><%= obrandCs.FItemList(i).FIsUpchebeasong %></td>
    <td ></td>
</tr>
<% next %>
<% end if %>
</table>

<script language='javascript'>
function chkComp(bool){
    document.frm.yyyy1.disabled = !(bool);
    document.frm.mm1.disabled = !(bool);
    document.frm.dd1.disabled = !(bool);

    document.frm.yyyy2.disabled = !(bool);
    document.frm.mm2.disabled = !(bool);
    document.frm.dd2.disabled = !(bool);

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
