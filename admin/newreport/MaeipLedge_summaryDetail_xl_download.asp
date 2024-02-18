<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"

Server.ScriptTimeOut = 60*30
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stockclass/monthlyMaeipLedgeCls.asp"-->
<!-- #include virtual="/lib/BarcodeFunction.asp"-->
<%
Dim CCADMIN : CCADMIN = C_ADMIN_AUTH
dim showItemDetailPopup : showItemDetailPopup = (left(now, 10) <= "2014-08-09") or CCADMIN


dim i, totalpage, totalCount
dim yyyy1, yyyymm1, makerid, showsuply, meaipTp, showShopid, showDiff
dim stockPlace, shopid
dim targetGbn, itemgubun
dim bPriceGbn, showItem
dim stype       '' S:���, J:����

dim page : page=1
dim PageSize : PageSize = 1000

stype       = requestCheckvar(request("stype"),10)
shopid    	= requestCheckvar(request("shopid"),32)
yyyy1       = requestCheckvar(request("yyyy1"),10)
stockPlace  = requestCheckvar(request("stockPlace"),10)
makerid     = requestCheckvar(request("makerid"),32)
showsuply   = requestCheckvar(request("showsuply"),10)
showShopid  = requestCheckvar(request("showShopid"),10)
meaipTp     = requestCheckvar(request("meaipTp"),10)
itemgubun   = requestCheckvar(request("itemgubun"),10)
targetGbn   = requestCheckvar(request("targetGbn"),10)
showDiff   	= requestCheckvar(request("showDiff"),10)
bPriceGbn   = requestCheckvar(request("bPriceGbn"),10)
showItem	= requestCheckvar(request("showItem"),2)

if (totalpage="") then totalpage=1
if (stype="") then stype="S"
if (stockPlace="") then stockPlace="L"
if (bPriceGbn="") then bPriceGbn = "V"
if (showsuply="") then showsuply = "on"
if (yyyy1="") then yyyy1 = "2012"

dim stockPlaceName
Select Case stockPlace
	Case "L" : stockPlaceName = "����"
	Case "S" : stockPlaceName = "����"
	Case "T" : stockPlaceName = "���"
	Case "O" : stockPlaceName = "�¶��θ�������"
	Case "N" : stockPlaceName = "�¶��θ�������_�����Ұ�"
	Case "F" : stockPlaceName = "������������"
	Case "A" : stockPlaceName = "�ΰŽ���������"
	Case "R" : stockPlaceName = "��Ż"
	Case "E" : stockPlaceName = "����"
End Select

Dim AdmPath : AdmPath = "/admin/newreport/xldwn/" & year(now)
Dim appPath : appPath = server.mappath(AdmPath) + "/"

Dim FileName: FileName = "YearMaeipLedge_" & yyyy1 & "_" & stockPlaceName & "_" & GetCurrentTimeFormat() &".csv"
dim fso, tFile, headLine, bodyLine
dim FTotCnt : FTotCnt=0

response.write "<div style=""text-align:center;"">"
response.write "<h3 style=""margin:40px 0 20px 0;"">�ٿ�ε� ���� <strong>["&FileName&"]</strong> <span id=""statusHead"">���� ��...</span></h3>"
response.write "�ۼ��� �ð��� �ɸ��ϴ�. â�� �������ð� �ٿ�ε� â�� ��Ÿ�� ������ ��ٷ��ּ���.<br /><br />"
response.write "<div style=""width:400px; padding:1px; background-color:#EEE; margin:auto;""><div id=""statusBar"" style=""width:0px; height:9px; background-image: linear-gradient(to right, #fFf6f0, #E53C3C);""></div></div>"
response.write "<h4 id=""statusText""></h4>"
response.write "</div>"
response.flush

'-----------------------------------------------------------------------
'// ������� ����
Set fso = CreateObject("Scripting.FileSystemObject")
	If NOT fso.FolderExists(appPath) THEN
		fso.CreateFolder(appPath)
	END If
Set tFile = fso.CreateTextFile(appPath & FileName )

headLine = "���,����ID,�귣��ID,��ǰ����,��ǰ�ڵ�,�ɼ��ڵ�,���ڵ�,���Ա���,�����ġ,��������,��ǰ��,�ɼǸ�,"
headLine = headLine & "����������,�������ݾ�,���Լ���,���Աݾ�,�̵�����,�̵��ݾ�,�Ǹż���,�Ǹűݾ�,����������,�������ݾ�,"
headLine = headLine & "��Ÿ������,��Ÿ���ݾ�,�ν�������,�ν����ݾ�,CS������,CS���ݾ�,��������,�����ݾ�,�⸻������,�⸻���ݾ�"
tFile.WriteLine headLine


'//������ ���� �� LOOP
'-----------------------------------------------------------------------
dim oCMonthlyMaeipLedge
	
do until (totalpage < page)
	set oCMonthlyMaeipLedge = new CMonthlyMaeipLedge

	oCMonthlyMaeipLedge.FRectYYYY = yyyy1
	oCMonthlyMaeipLedge.FRectStockPlace = stockPlace
	oCMonthlyMaeipLedge.FRectShopid = shopid
	oCMonthlyMaeipLedge.FRectMakerid = makerid
	oCMonthlyMaeipLedge.FRectBySuplyPrice = CHKIIF(showsuply="on",1,0)
	oCMonthlyMaeipLedge.FRectMeaipTp = meaipTp
	oCMonthlyMaeipLedge.FRectItemgubun = itemgubun
	oCMonthlyMaeipLedge.FRectTargetGbn = targetGbn
	oCMonthlyMaeipLedge.FRectShowShopid = showShopid
	oCMonthlyMaeipLedge.FRectShowItem = showItem
	oCMonthlyMaeipLedge.FRectPriceGubun = bPriceGbn
	oCMonthlyMaeipLedge.FRectShowPurchaseType = "on"	''��������ǥ��

	oCMonthlyMaeipLedge.FRectShowDiff = showDiff

	oCMonthlyMaeipLedge.FPageSize = PageSize
	oCMonthlyMaeipLedge.FCurrPage = page

	if (stype="S") then
		oCMonthlyMaeipLedge.GetMaeipLedgeSUMSubDetail
	else
		oCMonthlyMaeipLedge.GetMaeipJungsanSumSubDetail
	end if

	'��ü ������ �ݿ�(loop��)
	if page=1 and totalpage<>oCMonthlyMaeipLedge.FtotalPage then
		totalpage = oCMonthlyMaeipLedge.FtotalPage
		totalCount = oCMonthlyMaeipLedge.FTotalCount
	end if

	if oCMonthlyMaeipLedge.FResultCount>0 then
		FTotCnt = FTotCnt + oCMonthlyMaeipLedge.FResultCount
		for i=0 to oCMonthlyMaeipLedge.FResultCount-1 
			'������� ���
			bodyLine = oCMonthlyMaeipLedge.FItemList(i).Fyyyymm & ","			'��󿬵�
			bodyLine = bodyLine & chkIIF(showShopid<>"",oCMonthlyMaeipLedge.FItemList(i).Fshopid,"") & ","		'����ID
			bodyLine = bodyLine & oCMonthlyMaeipLedge.FItemList(i).FMakerid & ","			'�귣��ID
			bodyLine = bodyLine & oCMonthlyMaeipLedge.FItemList(i).Fitemgubun & ","			'��ǰ����
			bodyLine = bodyLine & oCMonthlyMaeipLedge.FItemList(i).Fitemid & ","			'��ǰ�ڵ�
			bodyLine = bodyLine & oCMonthlyMaeipLedge.FItemList(i).Fitemoption & ","		'�ɼ��ڵ�
			bodyLine = bodyLine & BF_MakeTenBarcode(oCMonthlyMaeipLedge.FItemList(i).Fitemgubun, oCMonthlyMaeipLedge.FItemList(i).Fitemid, oCMonthlyMaeipLedge.FItemList(i).Fitemoption) & ","		'���ڵ�
			bodyLine = bodyLine & oCMonthlyMaeipLedge.FItemList(i).getMeaipTypeName & ","	'���Ա���
			bodyLine = bodyLine & oCMonthlyMaeipLedge.FItemList(i).FstockPlace & ","		'�����ġ
			bodyLine = bodyLine & oCMonthlyMaeipLedge.FItemList(i).FpurchaseTypeName & ","	'��������
			bodyLine = bodyLine & oCMonthlyMaeipLedge.FItemList(i).FitemName & ","			'��ǰ��
			bodyLine = bodyLine & oCMonthlyMaeipLedge.FItemList(i).FitemoptionName & ","	'�ɼǸ�
			bodyLine = bodyLine & oCMonthlyMaeipLedge.FItemList(i).FprevSysStockNo & ","	'����������
			bodyLine = bodyLine & oCMonthlyMaeipLedge.FItemList(i).FprevSysStockSum & ","	'�������ݾ�
			bodyLine = bodyLine & oCMonthlyMaeipLedge.FItemList(i).getIpgoNo & ","			'���Լ���
			bodyLine = bodyLine & oCMonthlyMaeipLedge.FItemList(i).getIpgoSum & ","			'���Աݾ�
			bodyLine = bodyLine & oCMonthlyMaeipLedge.FItemList(i).getMoveNo & ","			'�̵�����
			bodyLine = bodyLine & oCMonthlyMaeipLedge.FItemList(i).getMoveSum & ","			'�̵��ݾ�
			bodyLine = bodyLine & oCMonthlyMaeipLedge.FItemList(i).FSellNo & ","			'�Ǹż���
			bodyLine = bodyLine & oCMonthlyMaeipLedge.FItemList(i).FSellSum & ","			'�Ǹűݾ�
			bodyLine = bodyLine & oCMonthlyMaeipLedge.FItemList(i).FOffChulNo & ","			'����������
			bodyLine = bodyLine & oCMonthlyMaeipLedge.FItemList(i).FOffChulSum & ","		'�������ݾ�
			bodyLine = bodyLine & oCMonthlyMaeipLedge.FItemList(i).FEtcChulNo & ","			'��Ÿ������
			bodyLine = bodyLine & oCMonthlyMaeipLedge.FItemList(i).FEtcChulSum & ","		'��Ÿ���ݾ�
			bodyLine = bodyLine & oCMonthlyMaeipLedge.FItemList(i).FLossChulNo & ","		'�ν�������
			bodyLine = bodyLine & oCMonthlyMaeipLedge.FItemList(i).FLossChulSum & ","		'�ν����ݾ�
			bodyLine = bodyLine & oCMonthlyMaeipLedge.FItemList(i).FCsNo & ","				'CS������
			bodyLine = bodyLine & oCMonthlyMaeipLedge.FItemList(i).FCsSum & ","				'CS���ݾ�
			bodyLine = bodyLine & oCMonthlyMaeipLedge.FItemList(i).getTotErrNo & ","		'��������
			bodyLine = bodyLine & oCMonthlyMaeipLedge.FItemList(i).getTotErrSum & ","		'�����ݾ�
			bodyLine = bodyLine & oCMonthlyMaeipLedge.FItemList(i).FcurSysStockNo & ","		'�⸻������
			bodyLine = bodyLine & oCMonthlyMaeipLedge.FItemList(i).FcurSysStockSum & ""		'�⸻���ݾ�

			tFile.WriteLine bodyLine
		next
	end if
	set oCMonthlyMaeipLedge = Nothing
	page = page + 1

	if totalCount>0 then
		response.write "<script>" &_
				"document.getElementById(""statusText"").innerHTML=""" & formatNumber(totalCount,0) & "�� �� " & formatNumber(FTotCnt,0) & "�� [" & formatNumber(FTotCnt/totalCount*100,1) & "%]""; " &_
				"document.getElementById(""statusBar"").style.width=""" & round(FTotCnt/totalCount*400,0) & "px"";" &_
				"</script>"
		response.flush
	end if
Loop

tFile.Close
Set tFile = Nothing
Set fso = Nothing
response.write "<script>document.getElementById(""statusHead"").innerHTML=""���� �Ϸ�!"";</script>"
response.write "<p style=""text-align:center;"">" & FTotCnt&"�� ���� �Ϸ� ["&FileName&"]</p>"
%>
<script>
self.location.replace("<%=AdmPath&"/"&FileName%>");
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->
