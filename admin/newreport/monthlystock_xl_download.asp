<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"

Server.ScriptTimeOut = 60*20		' 20��
%>
<%
'###########################################################
' Description : ����ڻ�
' History : �̻� ����
'			2023.05.04 �ѿ�� ����(�˻������߰�, ��� ����� �����ؼ� ��ü���� �ڸ�Ʈ ����)
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/stockclass/monthlystockcls.asp"-->
<%

Const MaxPage   = 40
Const PageSize = 5000

Const isOnlySys = FALSE
Const isViewWonga =FALSE

dim yyyy1,mm1,isusing,sysorreal, research, newitem, vatyn, minusinc, bPriceGbn, i
dim mwgubun, buseo, itemgubun, stplace, purchasetype, showsuply, dtype, makerid, shopid, etcjungsantype, showDiff
dim brandUseYN
	yyyy1       = requestCheckvar(request("yyyy1"),10)
	mm1         = requestCheckvar(request("mm1"),10)
	isusing     = requestCheckvar(request("isusing"),10)
	sysorreal   = requestCheckvar(request("sysorreal"),10)
	research    = requestCheckvar(request("research"),10)
	newitem     = requestCheckvar(request("newitem"),10)
	mwgubun     = requestCheckvar(request("mwgubun"),10)
	vatyn       = requestCheckvar(request("vatyn"),10)
	minusinc   = requestCheckvar(request("minusinc"),10)
	bPriceGbn   = requestCheckvar(request("bPriceGbn"),10)
	buseo       = requestCheckvar(request("buseo"),10)
	itemgubun   = requestCheckvar(request("itemgubun"),10)
	purchasetype   = requestCheckvar(request("purchasetype"),10)
	stplace     = requestCheckvar(request("stplace"),10)
	showsuply   = requestCheckvar(request("showsuply"),10)
	dtype       = requestCheckvar(request("dtype"),10)
	makerid     = requestCheckvar(request("makerid"),32)
	shopid     = requestCheckvar(request("shopid"),32)
	etcjungsantype      = requestCheckvar(request("etcjungsantype"),10)
	showDiff      = requestCheckvar(request("showDiff"),10)
	brandUseYN      = requestCheckvar(request("brandUseYN"),10)

if (makerid<>"") then dtype=""
if (sysorreal="") then sysorreal="sys"
if (research="") and (bPriceGbn = "") then
    bPriceGbn="V"
end if
if (stplace="") then
    stplace="L"
	showDiff = "Y"
end if
if (research="") then
	if (itemgubun = "") then
		'itemgubun = "AA"
	end if
	if (buseo = "") then
		buseo = "3X"
	end if
end if

dim nowdate
if yyyy1="" then
	nowdate = dateserial(year(Now),month(now)-1,1)
	yyyy1 = Left(CStr(nowdate),4)
	mm1 = Mid(CStr(nowdate),6,2)
end if

dim totno, totbuy, subTotno, subTotbuy '', totavgBuy, offtotavgBuy
dim totPreno, totPrebuy     , subPreno, subPrebuy
dim totIpno,totIpBuy        , subIpno, subIpBuy
dim totLossno, totLossBuy   , subLossno, subLossBuy
dim totSellno, totSellBuy   , subSellno, subSellBuy
dim totOffChulno, totOffChulBuy  , subOffChulno, subOffChulBuy
dim totEtcChulno, totEtcChulBuy  , subEtcChulno, subEtcChulBuy
dim totCsChulno, totCsChulBuy    , subCsChulno, subCsChulBuy
dim iURL, iURLEtc, nBusiName, diffStock, diffStockPrc, diffStockW
DIM isGroupByBrand : isGroupByBrand = (dtype="mk")
Dim isItemList : isItemList = (makerid<>"")

dim totErrBadItemno, totErrBadItemBuy, subErrBadItemno, subErrBadItemBuy
dim totMoveItemno, totMoveItemBuy, subMoveItemno, subMoveItemBuy
dim totErrRealCheckno, totErrRealCheckBuy, subErrRealCheckno, subErrRealCheckBuy
dim totRealStockno, totRealStockBuy, subRealStockno, subRealStockBuy
dim totErrRealCheckBuyPlus, totErrRealCheckBuyMinus

'Dim AdmPath : AdmPath = "/admin/newreport/xldwn/" & Replace(yyyymm, "-", "")
Dim AdmPath : AdmPath = "/admin/newreport/xldwn/" & yyyy1 & mm1
Dim appPath : appPath = server.mappath(AdmPath) + "/"

Dim sNow, sY, sM, sD, sH, sMi, sS, sDateName
	sNow = now()
	sY= Year(sNow)
	sM = Format00(2,Month(sNow))
	sD = Format00(2,Day(sNow))
	sH = Format00(2,Hour(sNow))
	sMi = Format00(2,Minute(sNow))
	sS = Format00(2,Second(sNow))
	sDateName = sY&sM&sD&sH&sMi&sS

Dim FileName: FileName = "MonthlyStockAsset_"&sDateName&".csv"
dim fso, tFile

Function WriteMakeFile(tFile, arrList)
    Dim intLoop,iRow
    Dim bufstr, tmpPrice
    Dim itemid,deliverytype, deliv
    iRow = UBound(arrList,2)
    For intLoop=0 to iRow
    	bufstr = ""

		bufstr = "'" & arrList(1,intLoop) & "'"		' YYYY-MM
		bufstr = bufstr & "," & arrList(2,intLoop)		' �����ġ
		bufstr = bufstr & "," & arrList(3,intLoop)		' ��ǰ����
		bufstr = bufstr & "," & arrList(4,intLoop)		' ��ǰ�ڵ�
		bufstr = bufstr & "," & "'" & arrList(5,intLoop) & "'"		' �ɼ��ڵ�
        bufstr = bufstr & "," & arrList(44,intLoop)				'���ڵ�

		bufstr = bufstr & "," & arrList(6,intLoop)		' �μ�����

		bufstr = bufstr & "," & arrList(52,intLoop)				'�귣��ID
		bufstr = bufstr & "," & arrList(53,intLoop)				'��������

		bufstr = bufstr & "," & arrList(7,intLoop)				'�����Ա���
		bufstr = bufstr & "," & arrList(55,intLoop)				'ON���Ա���
		bufstr = bufstr & "," & arrList(8,intLoop)		' ��������
		bufstr = bufstr & "," & arrList(10,intLoop)		' �������(SYS)

		if (bPriceGbn = "P") then
			bufstr = bufstr & "," & arrList(11,intLoop)
			bufstr = bufstr & "," & arrList(12,intLoop)		' ����(�԰�)
			bufstr = bufstr & "," & arrList(13,intLoop)
			bufstr = bufstr & "," & arrList(56,intLoop)		' �̵�
			bufstr = bufstr & "," & arrList(57,intLoop)
			bufstr = bufstr & "," & arrList(14,intLoop)		' �Ǹ�
			bufstr = bufstr & "," & arrList(15,intLoop)
			' arrList(16,intLoop) ���������� arrList(17,intLoop)
			bufstr = bufstr & "," & arrList(16,intLoop)+(arrList(56,intLoop)*-1)		' ������1(����������+�̵�����)
			bufstr = bufstr & "," & arrList(17,intLoop)+(arrList(57,intLoop)*-1)
			bufstr = bufstr & "," & arrList(18,intLoop)-arrList(20,intLoop)		' ������2(������ ��Ÿ��� �ν��� ���ԵǾ� ����)
			bufstr = bufstr & "," & arrList(19,intLoop)-arrList(21,intLoop)
			bufstr = bufstr & "," & arrList(20,intLoop)		' �����Ÿ���(�ν����)
			bufstr = bufstr & "," & arrList(21,intLoop)
			bufstr = bufstr & "," & arrList(22,intLoop)		' CS���
			bufstr = bufstr & "," & arrList(23,intLoop)
			' ����( (�������(SYS)+����(�԰�)+�̵�+�Ǹ�+������1+������2+CS���+�����Ÿ���)-�⸻��� )
			bufstr = bufstr & "," & ((arrList(10,intLoop) + arrList(12,intLoop) + arrList(56,intLoop) + arrList(14,intLoop) + (arrList(16,intLoop)+(arrList(56,intLoop)*-1)) + (arrList(18,intLoop)-arrList(20,intLoop)) + arrList(22,intLoop) + arrList(20,intLoop))-arrList(24,intLoop))*-1
			bufstr = bufstr & "," & ((arrList(11,intLoop) + arrList(13,intLoop) + arrList(57,intLoop) + arrList(15,intLoop) + (arrList(17,intLoop)+(arrList(57,intLoop)*-1)) + (arrList(19,intLoop)-arrList(21,intLoop)) + arrList(23,intLoop) + arrList(21,intLoop))-arrList(25,intLoop))*-1
			bufstr = bufstr & "," & arrList(24,intLoop)		' �ý������
			bufstr = bufstr & "," & arrList(25,intLoop)
			bufstr = bufstr & "," & arrList(28,intLoop)		' ��������
			bufstr = bufstr & "," & arrList(29,intLoop)
			bufstr = bufstr & "," & arrList(30,intLoop)		' �ǻ����
			bufstr = bufstr & "," & arrList(31,intLoop)
			bufstr = bufstr & "," & arrList(59,intLoop)		' �����ҷ�
			bufstr = bufstr & "," & arrList(60,intLoop)
		else
			bufstr = bufstr & "," & arrList(32,intLoop)
			bufstr = bufstr & "," & arrList(12,intLoop)		' ����(�԰�)
			bufstr = bufstr & "," & arrList(33,intLoop)
			bufstr = bufstr & "," & arrList(56,intLoop)		' �̵�
			bufstr = bufstr & "," & arrList(58,intLoop)
			bufstr = bufstr & "," & arrList(14,intLoop)		' �Ǹ�
			bufstr = bufstr & "," & arrList(34,intLoop)
			' arrList(16,intLoop) ���������� arrList(35,intLoop)
			bufstr = bufstr & "," & arrList(16,intLoop)+(arrList(56,intLoop)*-1)		' ������1(����������+�̵�����)
			bufstr = bufstr & "," & arrList(35,intLoop)+(arrList(58,intLoop)*-1)
			bufstr = bufstr & "," & arrList(18,intLoop)-arrList(20,intLoop)		' ������2(������ ��Ÿ��� �ν��� ���ԵǾ� ����)
			bufstr = bufstr & "," & arrList(36,intLoop)-arrList(37,intLoop)
			bufstr = bufstr & "," & arrList(20,intLoop)		' �����Ÿ���(�ν����)
			bufstr = bufstr & "," & arrList(37,intLoop)
			bufstr = bufstr & "," & arrList(22,intLoop)		' CS���
			bufstr = bufstr & "," & arrList(38,intLoop)
			' ����( (�������(SYS)+����(�԰�)+�̵�+�Ǹ�+������1+������2+CS���+�����Ÿ���)-�⸻��� )
			bufstr = bufstr & "," & ((arrList(10,intLoop) + arrList(12,intLoop) + arrList(56,intLoop) + arrList(14,intLoop) + (arrList(16,intLoop)+(arrList(56,intLoop)*-1)) + (arrList(18,intLoop)-arrList(20,intLoop)) + arrList(22,intLoop) + arrList(20,intLoop))-arrList(24,intLoop))*-1
			bufstr = bufstr & "," & ((arrList(32,intLoop) + arrList(33,intLoop) + arrList(58,intLoop) + arrList(34,intLoop) + (arrList(35,intLoop)+(arrList(58,intLoop)*-1)) + (arrList(36,intLoop)-arrList(37,intLoop)) + arrList(38,intLoop) + arrList(37,intLoop))-arrList(39,intLoop))*-1
			bufstr = bufstr & "," & arrList(24,intLoop)		' �ý������
			bufstr = bufstr & "," & arrList(39,intLoop)
			bufstr = bufstr & "," & arrList(28,intLoop)		' ��������
			bufstr = bufstr & "," & arrList(41,intLoop)
			bufstr = bufstr & "," & arrList(30,intLoop)		' �ǻ����
			bufstr = bufstr & "," & arrList(42,intLoop)
			bufstr = bufstr & "," & arrList(59,intLoop)		' �����ҷ�
			bufstr = bufstr & "," & arrList(61,intLoop)
		end if

		bufstr = bufstr & "," & "'" & arrList(43,intLoop) & "'"		' �����԰��

        bufstr = bufstr & "," & arrList(45,intLoop)		' ��ī�װ��ڵ�
        bufstr = bufstr & "," & arrList(46,intLoop)		' ��ī�װ���
        bufstr = bufstr & "," & arrList(47,intLoop)		' �߰�ī�װ��ڵ�
        bufstr = bufstr & "," & arrList(48,intLoop)		' �߰�ī�װ���
        bufstr = bufstr & "," & arrList(49,intLoop)		' ������X
        bufstr = bufstr & "," & arrList(50,intLoop)		' ������Y
        bufstr = bufstr & "," & arrList(51,intLoop)		' ������Z
		bufstr = bufstr & "," & arrList(54,intLoop)		'�Ǹſ���

		if (bPriceGbn = "P") then
			bufstr = bufstr & "," & arrList(26,intLoop)		' �������
			bufstr = bufstr & "," & arrList(27,intLoop)
			bufstr = bufstr & "," & arrList(62,intLoop)		' ����ҷ�
			bufstr = bufstr & "," & arrList(63,intLoop)
		else
			bufstr = bufstr & "," & arrList(26,intLoop)		' �������
			bufstr = bufstr & "," & arrList(40,intLoop)
			bufstr = bufstr & "," & arrList(62,intLoop)		' ����ҷ�
			bufstr = bufstr & "," & arrList(64,intLoop)
		end if

        tFile.WriteLine bufstr
    Next
End function

Dim sqlStr
Dim FTotCnt, FTotPage, FCurrPage

''rw "������"
''response.end

sqlStr = " [db_datamart].[dbo].[sp_Ten_monthlystock_Asset_MakeEXL_Count] ('" & yyyy1 & "-" & mm1 & "','" & stplace & "','" & shopid & "','"&buseo&"','"&itemgubun&"','"&mwgubun&"','"&vatyn&"','"&purchasetype&"','"&CHKIIF(showsuply="on",1,0)&"','"&CHKIIF(dtype="mk",1,0)&"','"&etcjungsantype&"','" & brandUseYN & "','') "

response.write sqlStr & "<br>"
db3_rsget.Open sqlStr, db3_dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
db3_dbget.CommandTimeout = 60*10   ' 10��
IF Not (db3_rsget.EOF OR db3_rsget.BOF) THEN
	FTotCnt = db3_rsget(0)
END IF
db3_rsget.close

response.write "FTotCnt:" & FTotCnt & "<br>"

Dim ArrRows
Dim headLine
IF FTotCnt > 0 THEN
	FTotPage =  CInt(FTotCnt\PageSize)
	If (FTotCnt\PageSize) <> (FTotCnt/PageSize) Then
		FTotPage = FTotPage + 1
	End If
    IF (FTotPage>MaxPage) THEn FTotPage=MaxPage

    Set fso = CreateObject("Scripting.FileSystemObject")
		If NOT fso.FolderExists(appPath) THEN
			fso.CreateFolder(appPath)
		END If
	Set tFile = fso.CreateTextFile(appPath & FileName )

	headLine = "YYYY-MM,�����ġ,��ǰ����,��ǰ�ڵ�,�ɼ��ڵ�,���ڵ�,�μ�����,�귣��ID,��������,�����Ա���,ON���Ա���,��������,�������(SYS),,����,,�̵�,,�Ǹ�,,������1,,������2,,�����Ÿ���,,CS���,,����,,�ý������,,��������,,�ǻ����,,�����ҷ�,,�����԰��,��ī�װ��ڵ�,��ī�װ���,�߰�ī�װ��ڵ�,�߰�ī�װ���,������X,������Y,������Z,�Ǹſ���,�������,,����ҷ�,,"

	tFile.WriteLine headLine

    For i=0 to FTotPage-1
    	ArrRows = ""
		sqlStr ="[db_datamart].[dbo].[sp_Ten_monthlystock_Asset_MakeEXL_List] ('" & yyyy1 & "-" & mm1 & "','" & stplace & "','" & shopid & "','"&buseo&"','"&itemgubun&"','"&mwgubun&"','"&vatyn&"','"&purchasetype&"','"&CHKIIF(showsuply="on",1,0)&"','"&CHKIIF(dtype="mk",1,0)&"','"&etcjungsantype&"','" & brandUseYN & "',''," & (i+1) & "," & PageSize & ")"

		response.write sqlStr & "<br>"
		db3_rsget.Open sqlStr, db3_dbget, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
        db3_dbget.CommandTimeout = 60*10   ' 10��
        IF Not (db3_rsget.EOF OR db3_rsget.BOF) THEN
        	ArrRows = db3_rsget.getRows()
        END IF
        db3_rsget.close
       	CALL WriteMakeFile(tFile,ArrRows)
    NExt
    tFile.Close
	Set tFile = Nothing
	Set fso = Nothing
END IF

response.write FTotCnt&"�� ���� ["&FileName&"]"
IF FTotCnt > 0 THEN
    response.redirect AdmPath&"/"&FileName
end if
''response.end
''response.write appPath & FileName
%>
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
