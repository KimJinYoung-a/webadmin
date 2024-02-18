<%@ language=vbscript %>
<% option explicit %>
<% Response.CharSet = "EUC-KR" %>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description : 데이터분석
' History : 2016.01.29 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbAnalopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/function.asp"-->

<!-- #include virtual="/lib/classes/dataanalysis/dataanalysis_cls.asp"-->
<!-- #include virtual="/lib/classes/dataanalysis/dataanalysis_salesissue_cls.asp"-->
<%
dim startdate, enddate, i, j
	startdate = Request("startdate")
	enddate = Request("enddate")

if startdate="" then startdate=date
if enddate="" then enddate=date

dim osales
set osales = new cdataanalysis_salesissue
	osales.FPageSize = 5
	osales.FCurrPage = 1
	osales.frectisusing = "Y"
	osales.frectstartdate = startdate
	osales.frectenddate = enddate
	osales.getdataanalysis_salesissue_top()
%>
<table width="100%" valign="top" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr bgcolor="<%= adminColor("tabletop") %>">
	<td align="left" colspan=2>
		<b>영업이슈</b>
	</td>
</tr>
<% if osales.FresultCount>0 then %>
	<% for j=0 to osales.FresultCount-1 %>
	<tr bgcolor="#FFFFFF">
		<td align="left">
			<%= FormatDate(osales.FItemList(j).fstartdate,"00.00") %> ~ <%= FormatDate(osales.FItemList(j).fenddate,"00.00") %>
		</td>
		<td align="left">
			<%= chrbyte(osales.FItemList(j).ftitle,30,"Y") %>
		</td>
	</tr>
	<% next %>
<% else %>
	<tr bgcolor="#FFFFFF">
		<td align="left" colspan=2>
			<b>영업이슈 검색 결과가 없습니다.</b>
		</td>
	</tr>
<% end if %>
</table>

<% set osales = nothing %>
<!-- #include virtual="/lib/db/dbAnalclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->