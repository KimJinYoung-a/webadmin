<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/etc/yahooitemcls.asp"-->
<%
'' dbget.close()	:	response.End
'' ���̸���(���ݺ� ����Ʈ)���� �ܾ, (����Ʈ�� ���޵Ǿ�����;�������:http://www.mm.co.kr/shop_admin/reg/shop_plist.asp?menu=03)

dim nowdate
dim adate,bdate
dim fso, FileName,tFile,appPath
dim readtextfile

appPath = server.mappath("/admin/etc/nate/") + "\"
FileName = "nateitem.txt"

nowdate = now()
adate = CDate(Left(nowdate,10) + " 09:00:00")
bdate = CDate(Left(nowdate,10) + " 17:00:00")

dim sqlStr,ref
ref = Left(request.ServerVariables("REMOTE_ADDR"),250)

sqlStr = "insert into [db_temp].[dbo].tbl_nate_scraplog"
sqlStr = sqlStr + " (ref) values('" + "NAT1-" + ref + "')"
dbget.execute sqlStr

if ((nowdate>adate) and (nowdate<bdate)) then
    '���� 9�� ~ ���� 6�ÿ��� �ٽ� ������ ���������� ������ �ڷ� �״�� ���
    '������Ʈ �ֱ� 08, 12, 15, 18��(�� 4��)
    response.redirect "/admin/etc/nate/" & FileName
    dbget.close()	:	response.End
end if

dim oNate, buf
dim totalpage, totalcount
dim ix, j

'// ���� ��Ʈ�� ȣ��
Set fso = CreateObject("Scripting.FileSystemObject")
Set tFile = fso.CreateTextFile(appPath & FileName )

'// ��ü ���������� �ѻ�ǰ�� ����
set oNate = new CYahooItemList
oNate.FPageSize = 300
oNate.FScrollCount = 100
oNate.GetNateItemCountDB3
	totalpage = oNate.FTotalPage
	totalcount = oNate.FTotalCount
set oNate = Nothing

buf = "<p>TOTAL:" & totalcount
tFile.WriteLine buf

'// ������ ����
for j=0 to totalpage - 1
	set oNate = new CYahooItemList
	oNate.FPageSize = 300
	oNate.FScrollCount = 100
	oNate.FTotalCount = totalpage
	oNate.FTotalPage = totalcount
	oNate.FCurrPage = j+1
	oNate.GetNateItemDB3  

	for ix=0 to oNate.FResultCount-1
		buf = "<p>"
		buf = buf & "tenbyten" & oNate.FItemList(ix).FItemId & "[^]"			'��ǰ�ڵ�(���θ�ID+��ǰ�ڵ�)
		buf = buf & oNate.FItemList(ix).GetModelname & "[^]"					'��ǰ��
		buf = buf & oNate.FItemList(ix).GetItemUrl & "[^]"						'��ǰ��ũ
		buf = buf & oNate.FItemList(ix).GetPrice & "[^]"						'�ǸŰ�
		buf = buf & oNate.FItemList(ix).getNatePath & "[^]"						'��ǰ�з�(ī�װ�)
		buf = buf & oNate.FItemList(ix).Getmakername & "[^]"					'������(�귣��)
		buf = buf & oNate.FItemList(ix).GetImageUrl & "[^]"						'��ǰ�̹���
		buf = buf & oNate.FItemList(ix).GetDeliverPay & "��[^]"					'��۷�
		buf = buf & oNate.FItemList(ix).GetMMCouponStr & "[^]"					'���αݾ�/���� (����� ����)
		buf = buf & "[^][^]"

		if buf<>"" then
			tFile.WriteLine buf
		end if
	next

	set oNate = Nothing
next

tFile.Close

Set tFile = Nothing
Set fso = Nothing

sqlStr = "insert into [db_temp].[dbo].tbl_nate_scraplog"
sqlStr = sqlStr + " (ref) values('" + "NAT2-" + ref + "')"
dbget.execute sqlStr
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
<%
response.redirect "/admin/etc/nate/" & FileName
%>