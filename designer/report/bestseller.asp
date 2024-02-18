<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/lib/designerbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/report/reportcls.asp"-->
<%
'###############################################
' PageName : bestseller.asp
' Discription : 업체어드민 베스트셀러 통계
' History : 2008.07.01 허진원 : 없는 날짜입력을 재계산 하도록 수정
'###############################################

'response.write "잠시 점검중입니다."
'dbget.close()
'response.end

dim yyyy1,mm1,dd1,yyyy2,mm2,dd2
dim yyyymmdd1,yyymmdd2
dim nowdate,searchnextdate
dim orderserial,itemid,ojumun
dim topn,page
dim ix,iy,cknodate

yyyy1 = requestCheckVar(request("yyyy1"),20)
mm1 = requestCheckVar(request("mm1"),20)
dd1 = requestCheckVar(request("dd1"),20)
yyyy2 = requestCheckVar(request("yyyy2"),20)
mm2 = requestCheckVar(request("mm2"),20)
dd2 = requestCheckVar(request("dd2"),20)
orderserial = requestCheckVar(request("orderserial"),20)
itemid = requestCheckVar(request("itemid"),20)
topn = requestCheckVar(request("topn"),20)
cknodate = requestCheckVar(request("cknodate"),20)

if (yyyy1="") then
	nowdate = Left(CStr(now()),10)
	yyyy1 = Left(nowdate,4)
	mm1   = Mid(nowdate,6,2)
	dd1   = Mid(nowdate,9,2)

	yyyy2 = yyyy1
	mm2   = mm1
	dd2   = dd1
else
	nowdate = Left(CStr(DateSerial(yyyy1 , mm1 , dd1)),10)
	yyyy1 = Left(nowdate,4)
	mm1   = Mid(nowdate,6,2)
	dd1   = Mid(nowdate,9,2)
end if

searchnextdate = Left(CStr(DateAdd("d",DateSerial(yyyy2 , mm2 , dd2),1)),10)

if (topn="") then topn=100

if topn>1000 then topn=1000

set ojumun = new CJumunMaster

if cknodate="" then
	ojumun.FRectRegStart = yyyy1 + "-" + mm1 + "-" + dd1
	ojumun.FRectRegEnd = searchnextdate
end if

ojumun.FRectDesignerID = session("ssBctID")
ojumun.FPageSize = topn
ojumun.FRectOrderSerial = orderserial
ojumun.FCurrPage = page
ojumun.SearchJumunListBybestseller

%>
<script language='javascript'>
function ViewOrderDetail(itemid){


window.open("http://www.10x10.co.kr/shopping/category_prd.asp?itemid=" + itemid,"sample");


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
</script>

<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
	<input type="hidden" name="menupos" value="<%= request("menupos") %>">
   	<tr height="10" valign="bottom">
        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td background="/images/tbl_blue_round_02.gif"></td>
        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
        <td background="/images/tbl_blue_round_04.gif"></td>
        <td>
        	기간 : <% DrawDateBox yyyy1,yyyy2,mm1,mm2,dd1,dd2 %>
        	&nbsp;&nbsp;
			검색갯수 : <input type="text" name="topn" value="<%= topn %>" size="7" maxlength="6">
			&nbsp;&nbsp;
			검색결과 : 총 <font color="red"><% = ojumun.FResultCount %></font>개
        </td>
        <td align="right">
        	<a href="javascript:document.frm.submit();"><img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
        </td>
        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form>
</table>
<!-- 표 상단바 끝-->

<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="100">상품코드</td>
		<td>상품명</td>
		<td>옵션명</td>
		<td width="80">공급가</td>
		<td width="60">총수량</td>
		<td width="80">공급가합계</td>
	</tr>
	<% if ojumun.FResultCount<1 then %>
	<tr>
		<td colspan="12" align="center" bgcolor="#FFFFFF">[검색결과가 없습니다.]</td>
	</tr>
	<% else %>
		<% for ix=0 to ojumun.FResultCount -1 %>
	<%
	Dim sumprice,totalsumprice
	sumprice = ojumun.FMasterItemList(ix).FBuycash * ojumun.FMasterItemList(ix).FItemNo
	%>
		<% if ojumun.FMasterItemList(ix).IsAvailJumun then %>
		<tr align="center" class="a" bgcolor="#FFFFFF">
		<% else %>
		<tr align="center" class="gray" bgcolor="#FFFFFF">
		<% end if %>
			<td height="25"><a href="#" onclick="ViewOrderDetail(<%= ojumun.FMasterItemList(ix).FItemID %>)" class="zzz"><%= ojumun.FMasterItemList(ix).FItemID  %></a></td>
			<td><%= ojumun.FMasterItemList(ix).FItemName %></td>
			<% if (ojumun.FMasterItemList(ix).FItemOptionStr="") then %>
				<td>&nbsp;</td>
			<% else %>
				<td><%= ojumun.FMasterItemList(ix).FItemOptionStr %></td>
			<% end if %>
			<td align="right"><%= FormatNumber(ojumun.FMasterItemList(ix).FBuycash,0)  %></td>
			<td><%= ojumun.FMasterItemList(ix).FItemNo %></td>
			<td align="right"><%= FormatNumber(sumprice,0) %></td>
		</tr>
		<% totalsumprice =  totalsumprice + sumprice %>
		<% next %>
	<% end if %>
</table>

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="<%= adminColor("topbar") %>">
    <tr valign="bottom" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="right">현재 페이지 합계 금액 : <font color="red"><% =FormatNumber(totalsumprice,0) %></font>원</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="top" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- 표 하단바 끝-->

<!-- #include virtual="/designer/lib/designerbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->