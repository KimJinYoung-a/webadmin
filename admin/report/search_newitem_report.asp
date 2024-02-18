<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/report/reportcls.asp"-->
<%
dim yyyy1,mm1,dd1,yyyy2,mm2,dd2
dim yyyymmdd1,yyymmdd2
dim nextdateStr,searchnextdate
dim ojumun
dim designer,page
dim ix,iy

yyyy1 = request("yyyy1")
mm1 = request("mm1")
dd1 = request("dd1")
yyyy2 = request("yyyy2")
mm2 = request("mm2")
dd2 = request("dd2")
designer = request("designer")

if (yyyy1="") then yyyy1 = Cstr(Year(now()))
if (mm1="") then mm1 = Cstr(Month(now()))
if (dd1="") then dd1 = Cstr(day(now()))
if (yyyy2="") then yyyy2 = Cstr(Year(now()))
if (mm2="") then mm2 = Cstr(Month(now()))
if (dd2="") then dd2 = Cstr(day(now()))

searchnextdate = Left(CStr(DateAdd("d",Cdate(yyyy2 + "-" + mm2 + "-" + dd2),1)),10)


set ojumun = new CJumunMaster
ojumun.FRectRegStart = yyyy1 + "-" + mm1 + "-" + dd1
ojumun.FRectRegEnd = searchnextdate
ojumun.FRectDesignerID = designer
ojumun.FPageSize = 30
ojumun.FCurrPage = page
ojumun.SearchNewItemReport

Dim totalsum
totalsum = 0
%>
<script language='javascript'>
function ViewOrderDetail(itemid){


window.open("http://www.10x10.co.kr/street/designershop.asp?itemid=" + itemid,"sample");


}

function ViewUserInfo(frm){
	//var popwin;
    //popwin = window.open('','userinfo');
    frm.target = 'userinfo';
    frm.action="viewuserinfo.asp"
	frm.submit();

}

function NextPage(ipage){
	document.frm.page.value= ipage;
	document.frm.submit();
}

function ReSearch(ifrm){
	ifrm.submit();
}
</script>
<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#CCCCCC">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
	<tr>
		<td class="a" >
		기간 :
		<% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
		디자이너 :
		<% drawSelectBoxDesigner "designer",designer %>
		</td>
		<td class="a" align="right">
			<a href="javascript:ReSearch(frm);"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
		</td>
	</tr>
	</form>
</table><br>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="a">
<% if ojumun.FResultCount<1 then %>
<tr>
	<td align="center">[검색결과가 없습니다.]</td>
</tr>
<% else %>
<tr>
	<% for ix=0 to ojumun.FResultCount -1 %>
<%
totalsum = totalsum + ojumun.FMasterItemList(ix).FTcnt
%>

		 <td width="120" align="center">
				<table border="0" cellpadding="0" cellspacing="1" bgcolor="#808080" width="100%" class="a">
				<tr bgcolor="#FFFFFF">
					<td><%= ojumun.FMasterItemList(ix).FCateName %></td>
				</tr>
				<tr bgcolor="#FFFFFF">
					<td><%= ojumun.FMasterItemList(ix).FTcnt %></td>
				</tr>
				</table>
		 </td>
	<% next %>
		 <td width="120" align="center">
				<table border="0" cellpadding="0" cellspacing="1" bgcolor="#808080"  width="100%" class="a">
				<tr bgcolor="#FFFFFF">
					<td>총계</td>
				</tr>
				<tr bgcolor="#FFFFFF">
					<td><%= totalsum %></td>
				</tr>
				</table>
		 </td>
	</tr>
<% end if %>
</table>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->