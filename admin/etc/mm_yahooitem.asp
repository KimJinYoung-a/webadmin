<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/etc/yahooitemcls.asp"-->
<%
''��� ���ϴµ�.?
''���� ���� ���..
'' MM (���ݺ� ����Ʈ) ���� �ܾ 1�ð� ����, (���Ŀ� ���޵Ǿ�����)
'' http://www.mm.co.kr/shop_admin/reg/dblist_01.txt
'' ��ǰ�ڵ�[^]��ǰ��[^]��ǰ��ũ[^]��ǰ����[^]��ǰ�з�[^]������[^]�̹���URL[^]��۷�[^]��������[^]����1[^]����2[^]����3
'' IIS 6.0 ���� ������ �������� response.flush �߰� or C:\Windows\System32\Inetsrv\Metabase.xml ����  AspBufferingLimit �� ����


dim nowdate
dim adate,bdate
nowdate = now()
adate = CDate(Left(nowdate,10) + " 09:00:00")
bdate = CDate(Left(nowdate,10) + " 23:59:59")

dim sqlStr,ref
ref = Left(request.ServerVariables("REMOTE_ADDR"),250)

sqlStr = "insert into [db_temp].[dbo].tbl_nate_scraplog"
sqlStr = sqlStr + " (ref) values('" + "YMM1-" + ref + "')"
dbget.execute sqlStr

dbget.close()	:	response.End


if ((nowdate>adate) and (nowdate<bdate)) then
    'dbget.close()	:	response.End
end if

dim oyahoo
dim page
page = request("page")
if page="" then page=1

dim ix

set oyahoo = new CYahooItemList
oyahoo.FPageSize = 30000
oyahoo.FScrollCount = 100
oyahoo.FCurrPage = page
oyahoo.GetYahooItemDB3  
%>
<p>Total : <%= oyahoo.FtotalCount & VbCrlf %>
<% for ix=0 to oyahoo.FResultCount-1 %><p><%= oyahoo.FItemList(ix).FItemId %>[^]<%= oyahoo.FItemList(ix).GetModelname %>[^]<%= oyahoo.FItemList(ix).GetItemUrl %>[^]<%= oyahoo.FItemList(ix).GetPrice %>[^]<%= Replace(oyahoo.FItemList(ix).FNmLarge,"/","") %>/<%= Replace(oyahoo.FItemList(ix).FNmMid,"/","") %>/<%= Replace(oyahoo.FItemList(ix).FNmSmall,"/","") %>[^]<%= oyahoo.FItemList(ix).Getmakername %>[^]<%= oyahoo.FItemList(ix).GetImageUrl %>[^]<%= oyahoo.FItemList(ix).GetDeliverPay %>[^]<%= oyahoo.FItemList(ix).GetMMCouponStr %>[^][^][^]
<% if ix mod 10000=0 then response.flush %>
<% next %>
<%
set oyahoo = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->