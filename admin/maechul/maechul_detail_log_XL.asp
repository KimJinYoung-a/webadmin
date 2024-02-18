<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"

Server.ScriptTimeOut = 60*5
%>
<%
'###########################################################
' Description : ���޸� Ŭ����
' Hieditor : 2011.04.22 �̻� ����
'			 2020.02.04 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%

Const MaxPage   = 40
Const PageSize = 5000  ''�Ǽ� ����.. (5000=>1000), ����ī�װ� �� => ���ν��� ����

dim makerid, startDate, endDate, vatinclude, mwdiv
dim yyyymm, Dategbn, indexmSqlStr, indexdSqlStr
Dategbn = requestCheckvar(request("Dategbn"),32)
makerid = request("makerid")
startDate = request("startDate")
endDate = request("endDate")
vatinclude = request("vatinclude")
mwdiv = request("mwdiv")

yyyymm = Left(startDate, 7)

if Dategbn="" then Dategbn="chulgoDate"

Dim AdmPath : AdmPath = "/admin/newreport/xldwn/" & Replace(yyyymm, "-", "")
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

Dim FileName: FileName = "maechul_detail_log_"&sDateName&".csv"
dim fso, tFile

function GetActDivCodeName(actDivCode)
	Select Case actDivCode
		Case "A"
			GetActDivCodeName = "���ֹ�"
		Case "C"
			GetActDivCodeName = "����ֹ�"
		Case "H"
			GetActDivCodeName = "��ǰ����"
		Case "E"
			GetActDivCodeName = "��ȯ�ֹ�"
		Case "M"
			GetActDivCodeName = "��ǰ�ֹ�"
		Case "CC"
			GetActDivCodeName = "�������ȭ"
		Case "HH"
			GetActDivCodeName = "��ǰ�������"
		Case "EE"
			GetActDivCodeName = "��ȯ���"
		Case "MM"
			GetActDivCodeName = "��ǰ���"
		Case Else
			GetActDivCodeName = actDivCode
	End Select
end function

function GetFullOrderSerial(orderserial, suborderserial)
	GetFullOrderSerial = orderserial & "-" & Format00(3, suborderserial)
end function

function GetVatIncludeName(vatinclude)
	Select Case vatinclude
		Case "N"
			GetVatIncludeName = "�鼼"
		Case Else
			GetVatIncludeName = "����"
	End Select
end function

function GetOMWdivName(omwdiv, itemid)
	if (CStr(itemid) = "0") then
		if (omwdiv="UU") then
			GetOMWdivName = "����"
		elseif (omwdiv="TT") then
			GetOMWdivName = "�ٹ�"
		else
			GetOMWdivName = omwdiv
		end if
	else
		Select Case omwdiv
			Case "M"
				GetOMWdivName = "����"
			Case "W"
				GetOMWdivName = "��Ź"
			Case "U"
				GetOMWdivName = "��ü"

			Case "B000"
				GetOMWdivName = "������"
			Case "B011"
				GetOMWdivName = "��Ź�Ǹ�"
			Case "B012"
				GetOMWdivName = "��ü��Ź"
			Case "B013"
				GetOMWdivName = "�����Ź"
			Case "B021"
				GetOMWdivName = "��������"
			Case "B022"
				GetOMWdivName = "�������"
			Case "B023"
				GetOMWdivName = "����������"
			Case "B031"
				GetOMWdivName = "������"
			Case "B032"
				GetOMWdivName = "���͸���"
			Case "B999"
				GetOMWdivName = "��Ÿ����"
			Case "PP"
				GetOMWdivName = "����"
			Case Else
				GetOMWdivName = omwdiv
		End Select
	end if
end function

Function WriteMakeFile(tFile, arrList)
    Dim intLoop,iRow
    Dim bufstr, tmpPrice
    Dim itemid,deliverytype, deliv, itemname, itemoptionname
    iRow = UBound(arrList,2)
    For intLoop=0 to iRow
    	bufstr = ""

		bufstr = "" & GetActDivCodeName(arrList(0,intLoop)) & ""
		bufstr = bufstr & "," & arrList(1,intLoop)
		bufstr = bufstr & "," & GetFullOrderSerial(arrList(2,intLoop), arrList(3,intLoop))
		bufstr = bufstr & "," & arrList(4,intLoop)
		bufstr = bufstr & "," & Left(arrList(5,intLoop), 10)
		bufstr = bufstr & "," & Left(arrList(6,intLoop), 10)
		bufstr = bufstr & "," & GetVatIncludeName(arrList(7,intLoop))
		bufstr = bufstr & "," & arrList(8,intLoop)

		bufstr = bufstr & "," & GetOMWdivName(arrList(9,intLoop), arrList(11,intLoop))
		bufstr = bufstr & "," & arrList(10,intLoop)
		bufstr = bufstr & "," & arrList(11,intLoop)
		bufstr = bufstr & ",'" & arrList(12,intLoop)

		itemname = db2html(arrList(13,intLoop))
		itemoptionname = db2html(arrList(14,intLoop))
		if (itemoptionname <> "") then
			itemname = itemname & "[" & itemoptionname & "]"
		end if
		itemname = Replace(itemname, Chr(34), "")
		itemname = Chr(34) & itemname & Chr(34)


		bufstr = bufstr & "," & itemname
		bufstr = bufstr & "," & arrList(15,intLoop)
		bufstr = bufstr & "," & arrList(16,intLoop)
		bufstr = bufstr & "," & arrList(17,intLoop)
		bufstr = bufstr & "," & arrList(18,intLoop)
		bufstr = bufstr & "," & arrList(19,intLoop)
		bufstr = bufstr & "," & arrList(20,intLoop)
		bufstr = bufstr & "," & arrList(21,intLoop)
		bufstr = bufstr & "," & arrList(22,intLoop)
		bufstr = bufstr & "," & arrList(23,intLoop)
		bufstr = bufstr & "," & arrList(24,intLoop)
		bufstr = bufstr & "," & arrList(25,intLoop)
		bufstr = bufstr & "," & arrList(26,intLoop)
		bufstr = bufstr & "," & arrList(27,intLoop)
		bufstr = bufstr & "," & arrList(28,intLoop)
		bufstr = bufstr & "," & arrList(29,intLoop)
		bufstr = bufstr & "," & arrList(30,intLoop)

        tFile.WriteLine bufstr
    Next
End function

dim sqlStr, addSqlStr
dim FTotCnt, FTotPage

indexmSqlStr = ""
indexdSqlStr = ""
if Dategbn="ActDate" then
	indexmSqlStr = indexmSqlStr + " with (NOLOCK,index(IX_tbl_order_master_log_actDate))"
elseif Dategbn="chulgoDate" then
	indexdSqlStr = indexdSqlStr + " with (NOLOCK,index(IX_tbl_order_detail_log_beasongdate))"
elseif Dategbn="jFixedDt" then
	indexdSqlStr = indexdSqlStr + " with (NOLOCK)"
else
	indexmSqlStr = indexmSqlStr + " with (NOLOCK,index(IX_tbl_order_master_log_ipkumdate))"
end if
if (application("Svr_Info")="Dev") then indexmSqlStr=" "
if (application("Svr_Info")="Dev") then indexdSqlStr=" "

addSqlStr = ""
' ����Ȯ������
if Dategbn="jFixedDt" Then
	addSqlStr = addSqlStr + " and d.DTLjFixedDt>='" + CStr(startDate) + "'"
	addSqlStr = addSqlStr + " and d.DTLjFixedDt<'" + CStr(endDate) + "'"

' ��������
elseif Dategbn="ActDate" Then
	addSqlStr = addSqlStr + " and m.actDate>='" + CStr(startDate) + "'"
	addSqlStr = addSqlStr + " and m.actDate<'" + CStr(endDate) + "'"

' ����������
elseif Dategbn="orgPay" Then
	addSqlStr = addSqlStr + " and m.ipkumdate>='" + CStr(startDate) + "'"
	addSqlStr = addSqlStr + " and m.ipkumdate<'" + CStr(endDate) + "'"


' �������
else
	addSqlStr = addSqlStr + " and d.beasongdate>='" + CStr(startDate) + "'"
	addSqlStr = addSqlStr + " and d.beasongdate<'" + CStr(endDate) + "'"
end if

addSqlStr = addSqlStr + " and d.vatinclude='" + vatinclude + "'"
if mwdiv="M" or mwdiv="W" or mwdiv="U" then
	addSqlStr = addSqlStr + " and d.itemid<>0 and d.omwdiv='" + mwdiv + "'"
elseif mwdiv="TT" then
	addSqlStr = addSqlStr + " and d.itemid=0 and left(d.itemoption, 1)<>'9'"
elseif mwdiv="UU" then
	addSqlStr = addSqlStr + " and d.itemid=0 and left(d.itemoption, 1)='9'"
elseif (Len(mwdiv) = 4) then
	addSqlStr = addSqlStr + " and d.omwdiv='" + mwdiv + "' "
end if
addSqlStr = addSqlStr + " and d.makerid='" + makerid + "'"

sqlStr = "select count(*) as cnt , CEILING(CAST(Count(*) AS FLOAT)/" + CStr(PageSize) + ") as totPg "
sqlStr = sqlStr + " from "
sqlStr = sqlStr + "		db_datamart.dbo.tbl_order_master_log m " & indexmSqlStr
sqlStr = sqlStr + "		join db_datamart.dbo.tbl_order_detail_log d " & indexdSqlStr
sqlStr = sqlStr + "		on "
sqlStr = sqlStr + "			1 = 1 "
sqlStr = sqlStr + "			and m.orderserial = d.orderserial "
sqlStr = sqlStr + "			and m.suborderserial = d.suborderserial "
sqlStr = sqlStr + " where "
sqlStr = sqlStr + "		1 = 1 "
sqlStr = sqlStr + addSqlStr

db3_rsget.CursorLocation = adUseClient
db3_rsget.Open sqlStr, db3_dbget, adOpenForwardOnly, adLockReadOnly'', adCmdStoredProc
IF Not (db3_rsget.EOF OR db3_rsget.BOF) THEN
	FTotCnt = db3_rsget(0)
END IF
db3_rsget.close

Dim i, ArrRows
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

	headLine = "����,����ó,�ֹ���ȣ,���ֹ���ȣ,��������,������(ó����),��������,��ǰ�ͼ�,���Ա���,�귣��,��ǰ�ڵ�,�ɼ��ڵ�,��ǰ��[�ɼǸ�],����,�Һ��ڰ��հ�,�ǸŰ�(���ΰ�),��ǰ�������밡,��������,��������,��ۺ�����,��Ÿ����(�þ�),�����Ѿ�,��ü�����,ȸ�����,�����,������,���Ÿ��ϸ���,�����,��ո��԰�,���"

	tFile.WriteLine headLine

    For i=0 to FTotPage-1
    	ArrRows = ""

		sqlStr = "select top " + CStr(PageSize*(i+1)) + " "
		sqlStr = sqlStr + " m.actDivCode "
		sqlStr = sqlStr + " , m.sitename "
		sqlStr = sqlStr + " , d.orderserial, d.suborderserial "
		sqlStr = sqlStr + " , (select top 1 linkorderserial from [db_order].[dbo].[tbl_order_master] o where o.orderserial = m.orderserial) as orgorderserial "
		sqlStr = sqlStr + " , m.ipkumdate "
		sqlStr = sqlStr + " , m.actDate "
		sqlStr = sqlStr + " , d.vatinclude "
		sqlStr = sqlStr + " , IsNull(m.targetGbn, 'ON') as targetGbn "

		sqlStr = sqlStr + " , d.omwdiv "
		sqlStr = sqlStr + " , d.makerid "
		sqlStr = sqlStr + " , d.itemid "
		sqlStr = sqlStr + " , d.itemoption "
		sqlStr = sqlStr + " , d.itemname, d.itemoptionname "			'// 14
		sqlStr = sqlStr + " , d.itemno "
		sqlStr = sqlStr + " , d.orgitemcost*d.itemno "
		sqlStr = sqlStr + " , d.itemcostCouponNotApplied*d.itemno "
		sqlStr = sqlStr + " , d.itemcost*d.itemno "

		sqlStr = sqlStr + " , (case when d.itemid <> 0 then (d.itemcost - d.reducedPrice)*d.itemno else 0 end) - d.anbunCouponPriceDetailSUM - allAtDiscount "
		sqlStr = sqlStr + " , d.anbunCouponPriceDetailSUM "
		sqlStr = sqlStr + " , (case when d.itemid = 0 then (d.itemcost - d.reducedPrice)*d.itemno else 0 end) "
		sqlStr = sqlStr + " , d.allAtDiscount "

		sqlStr = sqlStr + " , d.reducedPrice*d.itemno "
		sqlStr = sqlStr + " , d.upcheJungsanCash*d.itemno "
		sqlStr = sqlStr + " , (d.reducedPrice - d.upcheJungsanCash)*d.itemno "
		sqlStr = sqlStr + " , d.beasongdate "
		sqlStr = sqlStr + " , d.DTLjFixedDt"
		sqlStr = sqlStr + " , d.mileage*d.itemno "
		sqlStr = sqlStr + " , '' "
		sqlStr = sqlStr + " , IsNull((case "
		sqlStr = sqlStr + " 	when d.omwdiv in ('M', 'B031') then s.avgipgoPrice*d.itemno "
		sqlStr = sqlStr + "     else 0 end),0) as avgipgoPrice "
		sqlStr = sqlStr + " , '' "
		sqlStr = sqlStr + " from "
		sqlStr = sqlStr + " 	db_datamart.dbo.tbl_order_master_log m "
		sqlStr = sqlStr + " 	join db_datamart.dbo.tbl_order_detail_log d " & indexdSqlStr
		sqlStr = sqlStr + " 	on "
		sqlStr = sqlStr + " 		1 = 1 "
		sqlStr = sqlStr + " 		and m.orderserial = d.orderserial "
		sqlStr = sqlStr + " 		and m.suborderserial = d.suborderserial "
		sqlStr = sqlStr + "		Left Join db_summary.dbo.tbl_monthly_accumulated_logisstock_summary as s with(noLock) "
		sqlStr = sqlStr + "		on s.yyyymm=convert(varchar(7),m.actDate,21) "
		sqlStr = sqlStr + "			and s.itemgubun=d.itemgubun "
		sqlStr = sqlStr + "			and s.itemid=d.itemid "
		sqlStr = sqlStr + "			and s.itemoption=d.itemoption "
		sqlStr = sqlStr + " where "
		sqlStr = sqlStr + " 	1 = 1 "
		sqlStr = sqlStr + addSqlStr
    	sqlStr = sqlStr + " order by m.actDate desc, m.orderserial, m.suborderserial, d.itemgubun, d.itemid, d.itemoption"

		'response.write sqlStr & "<Br>"
		db3_rsget.CursorLocation = adUseClient
	    db3_rsget.pagesize = PageSize
		db3_rsget.Open sqlStr,db3_dbget,adOpenForwardOnly, adLockReadOnly
		db3_rsget.absolutepage = (i+1)

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
response.redirect AdmPath&"/"&FileName

%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
