<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/db2open.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/yahooitemcls.asp"-->
<%
'' dbget.close()	:	response.End
''서팀 문의 요망..
'' BB (가격비교 사이트) 에서 긁어감 1시간 단위, (야후와 제휴되어있음)

dim oyahoo
dim page
page = request("page")
if page="" then page=1

dim ix

set oyahoo = new CYahooItemList
oyahoo.FPageSize = 300
oyahoo.FScrollCount = 100
oyahoo.FCurrPage = page
oyahoo.GetYahooItemDB2
%>
<HTML>
<HEAD>
<TITLE>상품리스트 형식</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ks_c_5601-1987">
<META http-equiv=Cache-Control content=no-cache>
<META http-equiv=Expires content=0>
<META http-equiv=Pragma content=no-cache>
</HEAD>
<BODY>
<P>Total : <%= oyahoo.FtotalCount %>
<P>Serial : 200710231300
<% for ix=0 to oyahoo.FResultCount-1 %>
<p><%= oyahoo.FItemList(ix).FItemId %>^<%= oyahoo.FItemList(ix).FNmLarge %>^<%= oyahoo.FItemList(ix).FNmMid %>^<%= oyahoo.FItemList(ix).FNmSmall %>^<%= oyahoo.FItemList(ix).Getmakername %>^<%= oyahoo.FItemList(ix).GetModelname %>^<%= oyahoo.FItemList(ix).GetItemUrl %>^<%= oyahoo.FItemList(ix).GetImageUrl %>^<%= oyahoo.FItemList(ix).GetPrice %>^
<% next %>
<p>
	<% for ix=0 + oyahoo.StarScrollPage to oyahoo.FScrollCount + oyahoo.StarScrollPage - 1 %>
		<% if ix > oyahoo.FTotalpage then Exit for %>
		<a href="http://webadmin.10x10.co.kr/admin/etc/yahooitem.asp?page=<%= ix %>"><%= ix %></a>
	<% next %>
</BODY>
</HTML>
<%
set oyahoo = Nothing
%>
<!-- #include virtual="/lib/db/db2close.asp" -->