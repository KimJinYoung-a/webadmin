<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/etc/nateitemcls.asp"-->
<%
'' dbget.close()	:	response.End
'' BB(���ݺ� ����Ʈ)���� �ܾ (2011-09-20 ���̸������� BB�� ��ȯ)

dim page
page = request("page")
if page="" then page=1

dim sqlStr,ref
ref = Left(request.ServerVariables("REMOTE_ADDR"),250)

sqlStr = "insert into [db_temp].[dbo].tbl_nate_scraplog"
sqlStr = sqlStr + " (ref) values('" + "NAT3-" + ref + "')"
dbget.execute sqlStr

dim oNate, buf
dim totalpage, totalcount
dim ix

set oNate = new CNateItemList
oNate.FPageSize = 500
oNate.FScrollCount = 100
oNate.FTotalCount = totalpage
oNate.FTotalPage = totalcount
oNate.FCurrPage = page
oNate.GetNateItemDB3  

totalpage = oNate.FTotalPage
totalcount = oNate.FTotalCount

buf = "<<<total>>>" & vbCrLf
buf = buf & "	<<<�ѻ�ǰ��>>>" & formatNumber(totalcount,0) & vbCrLf
buf = buf & "	<<<����������>>>" & left(GetCurrentTimeFormat,12) & vbCrLf
'''buf = buf & "	<<<����/�߰���ǰ��>>>0" & vbCrLf
buf = buf & "<<</total>>>" & vbCrLf & vbCrLf

Response.Write buf

for ix=0 to oNate.FResultCount-1
	buf = "<<<product>>>" & vbCrLf
	buf = buf & "	<<<��ǰ���̵�>>>" & oNate.FItemList(ix).FItemId & vbCrLf
	buf = buf & "	<<<��ǰ��>>>" & oNate.FItemList(ix).GetModelname & vbCrLf
	buf = buf & "	<<<��ǰ�з���>>>" & oNate.FItemList(ix).getNateBBPath & vbCrLf							'��ǰ�з�(ī�װ�)
	buf = buf & "	<<<������>>>" & oNate.FItemList(ix).FitemMaker & vbCrLf
	buf = buf & "	<<<�����>>>" & vbCrLf
	buf = buf & "	<<<�귣��>>>" & oNate.FItemList(ix).Getmakername & vbCrLf
	buf = buf & "	<<<������>>>" & oNate.FItemList(ix).FsourceArea & vbCrLf
	buf = buf & "	<<<��ǰURL>>>" & Replace(oNate.FItemList(ix).GetItemUrl,"http://","") & vbCrLf			'��ǰ��ũ
	buf = buf & "	<<<��ǰ�̹���URL>>>" & Replace(oNate.FItemList(ix).GetListImageUrl,"http://","") & vbCrLf	'��ǰ�̹���(���:120)
	buf = buf & "	<<<��ǰū�̹���URL>>>" & Replace(oNate.FItemList(ix).GetBasicImageUrl,"http://","") & vbCrLf	'��ǰ�̹���(�⺻:400)
	buf = buf & "	<<<�ǸŰ�>>>" & formatNumber(oNate.FItemList(ix).GetPrice,0) & vbCrLf
	buf = buf & "	<<<������>>>" & vbCrLf
	buf = buf & "	<<<��۷�>>>" & formatNumber(oNate.FItemList(ix).GetDeliverPay,0) & vbCrLf
	buf = buf & "	<<<��۱Ⱓ>>>" & vbCrLf
	buf = buf& "	<<<��������>>>" & oNate.FItemList(ix).GetMMCouponStr & vbCrLf			'���αݾ�/���� (����� ����)
	buf = buf & "	<<<������>>>" & formatNumber(oNate.FItemList(ix).Fmileage,0) & vbCrLf
	buf = buf & "	<<<�������Һ�>>>" & vbCrLf																'����ǰ�������Һ� �����ô� ����
	buf = buf & "	<<<�̺�Ʈ>>>" & vbCrLf

	buf = buf & "<<</product>>>" & vbCrLf & vbCrLf

	Response.Write buf
next

buf = ""
for ix=0 + oNate.StarScrollPage to oNate.FScrollCount + oNate.StarScrollPage - 1
	if ix > oNate.FTotalpage then Exit for
	buf = buf & "<a href='http://webadmin.10x10.co.kr/admin/etc/nateitem_bb.asp?page=" & ix & "'>" & ix & "</a>"
next

Response.Write buf

set oNate = Nothing

sqlStr = "insert into [db_temp].[dbo].tbl_nate_scraplog"
sqlStr = sqlStr + " (ref) values('" + "NAT4-" + ref + "')"
dbget.execute sqlStr
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
