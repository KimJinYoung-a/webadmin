<%@ language=vbscript %>
<% option explicit %>
<%
'#######################################################
'	History	:  2010.09.10 �̻� ����
'			   2011.06.14 �ѿ�� ����
'	Description : �ֹ� �������� ���� ���
'#######################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db_TPLOpen.asp" -->
<!-- #include virtual="/lib/db/dbTPLHelper.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/admin/etc/orderInput/incUploadFunction.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<%

Dim paramInfo, retParamInfo, RetErr, retErrStr, sqlStr ,iLine, iResult ,POS1,POS2,POS3, okCNT
Dim xlPosArr, ArrayLen, skipString, afile, aSheetName ,i,j ,xlRowALL
Dim uploadform, objfile, sDefaultPath, sFolderPath
Dim iML, sFile, sFilePath, xltype, iMaxLen, sUploadPath, orgFileName, maybeSheetName
Dim monthFolder : monthFolder = Replace(Left(CStr(now()),7),"-","")
dim partnerItemName, partnerOptionName, barcode
dim makerid, itemgubun, itemid, itemoption, itemname, optionname, orgprice, mainimage, listimage, smallimage
dim itemoptionname
dim tplcompanyid

IF (application("Svr_Info")	= "Dev") then
    Set uploadform = Server.CreateObject("TABS.Upload")	   '' - TEST : TABS.Upload
ELSE
    Set uploadform = Server.CreateObject("TABSUpload4.Upload")	''REAL : TABSUpload4.Upload
END IF

Set objfile	   = Server.CreateObject("Scripting.FileSystemObject")
sDefaultPath   = Server.MapPath("/admin/etc/orderInput/upFiles/")

uploadform.Start sDefaultPath '���ε���

iMaxLen 		= uploadform.Form("iML")	'�̹�������ũ��
xltype 			= uploadform.Form("xltype")
tplcompanyid	= uploadform.Form("tplcompanyid")

IF (fnChkFile(uploadform("sFile"), iMaxLen,"xls")) THEN	'����üũ

    '���� ����
    sFolderPath = sDefaultPath&"/tplorder/"
    IF NOT  objfile.FolderExists(sFolderPath) THEN
    	objfile.CreateFolder sFolderPath
    END IF

    sFolderPath = sDefaultPath&"/order/"&monthFolder&"/"
    IF NOT  objfile.FolderExists(sFolderPath) THEN
    	objfile.CreateFolder sFolderPath
    END IF

    '��������
	sFile = fnMakeFileName(uploadform("sFile"))
	sFilePath = sFolderPath&sFile
	sFilePath = uploadform("sFile").SaveAs(sFilePath, False)

	orgFileName = uploadform("sFile").FileName
	maybeSheetName = Replace(orgFileName,"."&uploadform("sFile").FileType,"")

END IF

Set objfile		= Nothing
Set uploadform = Nothing

''''			�Ʒ� ������� �Ѵ�.
''''			������ 0���� �����Ѵ�.
''''			������addr1,������addr2 �� �Ѱ��� �ʵ��϶��� �����ϰ� �������ش�.

''''             �ֹ���ȣ, �ֹ���, �Ա���, ���Ҽ���, �ֹ���ID, �ֹ���, �ֹ�����ȭ,�ֹ����޴���ȭ,�ֹ����̸���,
''''             ������,��������ȭ,�������ڵ���,������Zip,������addr1,������addr2
''''             ��ǰ�ڵ�, �ɼ��ڵ�, ����, �ǸŰ�, �Һ��ڰ�, �����, ��ǰ��, �ɼǸ�,
''''             ��ü��ǰ�ڵ�,��ü�ɼ��ڵ�,�ֹ�������Ű, �ֹ����ǻ���, ��ǰ�䱸����1, ��ǰ�䱸����2, ��ǰ�䱸����3
''''              �ɼǰ���, �ɼǰ��ް���. �����ڵ�, ���θ���, ��ǰ��2, �ɼǸ�2

''''			���θ� �ֹ���ȣ, ���� �ֹ���ȣ, ���θ���, �����θ�, ��ȭ��ȣ1,
''''			��ȭ��ȣ2, �����ȣ, �ּ�, ����(��), Ư�����,
''''			���ӱ���, ����, ��ǰ��1, ��ǰ��(Ȯ��), �ɼǸ�,
''''			�ɼǸ�(Ȯ��), �ɼǺ�Ī(���ڵ�), �ǸŰ�(����), EA(��ǰ)*����

if (xltype = "sabangnet") then
	'' ����
	xlPosArr = Array(0,-1,-1,-1,-1,-1,-1,-1,-1,    3,4,5,6,7,7,    -1,-1,8,17,17,-1,12,14,    -1,-1,1,9,-1,-1,-1    -1,-1, 16, 2, 13,15)
	ArrayLen = UBound(xlPosArr)
	skipString="���θ� �ֹ���ȣ"
	afile = sFilePath
	aSheetName = maybeSheetName
elseif (xltype = "default") then
	'' �⺻����
	'' �ֹ���ȣ	�ֹ�������Ű	���θ���	������	��������ȭ	�������ڵ���
	'' ������Zip	������addr1	������addr2	��ǰ��	�ɼǸ�	����
	'' �����ڵ�	�Һ��ڰ�	�ǸŰ�	�ֹ����ǻ���
	xlPosArr = Array(0,-1,-1,-1,-1,-1,-1,-1,-1,    3,4,5,6,7,7,    -1,-1,11,13,12,-1,9,10,    -1,-1,1,14,-1,-1,-1    -1,-1, 11, 2, -1,-1)
	ArrayLen = UBound(xlPosArr)
	skipString="�ֹ���ȣ"
	afile = sFilePath
	aSheetName = maybeSheetName
else
    response.write "<script>alert('��ϵ��� ���� �����Դϴ�. -"&xltype&"');</script>"
    response.end
end if


if (true) then
	'// skip
elseif (companyid = "toms") then

	if (SellSite="5") then

	    '' Ž�� - ������
	    xlPosArr = Array(0,2,42,39,6,5,9,10,11,13,16,17,14,15,15,22,-1,28,29,31,-1,24,25,-1,-1,-1,12,19,20,21)
	    ArrayLen = UBound(xlPosArr)
	    skipString="�ֹ���ȣ"
	    afile = sFilePath
	    aSheetName = "TOMS �ֹ�����"  '' sheet name Maybe filename in

	elseif (SellSite="4") then

	    ''�ٹ����� - �ε����߰�����
	    xlPosArr = Array(1,2,-1,-1,-1,3,4,5,6,7,8,9,10,11,12,15,-1,19,18,-1,-1,16,17,21,-1,0,13,20,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString="�Ϸù�ȣ"
	    afile = sFilePath
	    aSheetName = maybeSheetName  '' sheet name Maybe filename in
	else
	    response.write "<script>alert('���θ� �ڵ尡 �������� �ʾҽ��ϴ�. -"&SellSite&"');</script>"
	    response.end
	end if

elseif (companyid = "ithinkso") then
    if (SellSite="6") then
        ''�ϳ�����
        xlPosArr = Array(0,15,15,-1,-1,2,4,5,-1,3,4,5,16,16,16,1,-1,8,9,9,10,6,7,-1,-1,-1,17,-1,-1,-1,11,12)
	    ArrayLen = UBound(xlPosArr)
	    skipString="�ֹ���ȣ"
	    afile = sFilePath
	    aSheetName = maybeSheetName  '' sheet name Maybe filename in
    elseif (SellSite="10") then
        ''(��)�����������ڸ���
        xlPosArr = Array(2,1,1,-1,-1,13,19,20,-1,14,17,18,15,16,16,5,-1,10,12,12,11,6,8,-1,-1,-1,21,-1,-1,-1,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString="�ֹ���ȣ"
	    afile = sFilePath
	    aSheetName = maybeSheetName  '' sheet name Maybe filename in
    elseif (SellSite="11") then
        ''��������
        xlPosArr = Array(4,2,2,-1,-1,		14,20,21,-1,15,		18,19,16,17,17,		6,-1,11,13,13,		12,7,8,-1,-1,		-1,23,-1,-1,-1,-1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString="�ֹ���ȣ"
	    afile = sFilePath
	    aSheetName = maybeSheetName  '' sheet name Maybe filename in
    elseif (SellSite="7") then
		''''            �ֹ���ȣ, �ֹ���, �Ա���, ���Ҽ���, �ֹ���ID

		''''			�ֹ���, �ֹ�����ȭ,�ֹ����޴���ȭ,�ֹ����̸���,������

		''''			��������ȭ,�������ڵ���,������Zip,������addr1,������addr2

		''''            ��ǰ�ڵ�, �ɼ��ڵ�, ����, �ǸŰ�, �Һ��ڰ�

		''''			�����, ��ǰ��, �ɼǸ�,��ü��ǰ�ڵ�,��ü�ɼ��ڵ�

		''''			�ֹ�������Ű, �ֹ����ǻ���, ��ǰ�䱸����1, ��ǰ�䱸����2, ��ǰ�䱸����3

		''''            �ɼǰ���, �ɼǰ��ް���.
    	'/����ī��ô�					   .             .                       .                    .
    	'xlPosArr = Array(1,23,23,-1,-1,2,3,4,5, 6,7,8,9,10,11 ,16,-1,19,20,20,-1,17,18 ,16,-1,-1,12,-1,-1,-1, -1,-1)	'2011.10.17 �� �������
    	xlPosArr = Array(1,23,23,-1,-1,     2,3,4,-1, 5,     6,7,8,9,10 ,     15,-1,18,19,19,     -1,16,17 ,14,-1,     -1,11,-1,-1,-1,      -1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString="�ֹ���ȣ"
	    afile = sFilePath
	    aSheetName = maybeSheetName  '' sheet name Maybe filename in
    elseif (SellSite="8") then

    	''��Ʈ��
    	''1			2		3			4		5			6			7			8		9			10					11		12				13		14
    	''�ֹ���ȣ	�ֹ���	�޴»��	����	��ȭ��ȣ	�����ȭ	�����ȣ	�ּ�	���޻���	���̶�� ��ǰ�ڵ�	��ǰ��	�ɼ�(Ư¡����)	�ǸŰ�	�ֹ���
    	''xlPosArr = Array(1,14,14,-1,-1,2,5,6,-1, 3,5,6,7,8,15 ,10,-1,4,13,13,-1,11,12 ,10,-1,-1,9,-1,-1,-1, -1,-1)
    	xlPosArr = Array(0,13,13,-1,-1,1,4,5,-1, 2,4,5,6,7,7 ,9,-1,3,12,12,-1,10,11 ,9,-1,-1,8,-1,-1,-1, -1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString="�ֹ���ȣ"
	    afile = sFilePath
	    aSheetName = maybeSheetName  '' sheet name Maybe filename in
    elseif (SellSite="9") then
		''''            �ֹ���ȣ, �ֹ���, �Ա���, ���Ҽ���, �ֹ���ID

		''''			�ֹ���, �ֹ�����ȭ,�ֹ����޴���ȭ,�ֹ����̸���,������

		''''			��������ȭ,�������ڵ���,������Zip,������addr1,������addr2

		''''            ��ǰ�ڵ�, �ɼ��ڵ�, ����, �ǸŰ�, �Һ��ڰ�

		''''			�����, ��ǰ��, �ɼǸ�,��ü��ǰ�ڵ�,��ü�ɼ��ڵ�

		''''			�ֹ�������Ű, �ֹ����ǻ���, ��ǰ�䱸����1, ��ǰ�䱸����2, ��ǰ�䱸����3

		''''            �ɼǰ���, �ɼǰ��ް���.
    	'/�м��÷���					   .             .                       .                    .
    	'xlPosArr = Array(1,23,23,-1,-1,2,3,4,5, 6,7,8,9,10,11 ,16,-1,19,20,20,-1,17,18 ,16,-1,-1,12,-1,-1,-1, -1,-1)	'2011.10.17 �� �������
    	xlPosArr = Array(1,5,5,-1,-1,     27,26,25,-1, 2,     26,25,24,22,22 ,     10,-1,12,13,13,     -1,11,4 ,10,-1,     -1,28,-1,-1,-1,      -1,-1)
	    ArrayLen = UBound(xlPosArr)
	    skipString="�ֹ���ȣ"
	    afile = sFilePath
	    aSheetName = maybeSheetName  '' sheet name Maybe filename in
    else
	    response.write "<script>alert('���θ� �ڵ尡 �������� �ʾҽ��ϴ�. -"&SellSite&"');</script>"
	    response.end
	end if
else
    response.write "<script>alert('��ϵ��� ���� ��ü�Դϴ�. -"&companyid&"');</script>"
    response.end
end if

ReDim xlRow(ArrayLen)
rw "ArrayLen="&ArrayLen

dim ret : ret = fnGetXLFileArray(xlRowALL, afile, aSheetName, ArrayLen)

if (Not ret) or (Not IsArray(xlRowALL)) then
    response.write "<script>alert('������ �ùٸ��� �ʰų� ������ �����ϴ�. "&Replace(Err.Description,"'","")&"');</script>"

    if (Err.Description="�ܺ� ���̺� ������ �߸��Ǿ����ϴ�.") then
        response.write "<script>alert('�������� Save As Excel 97 -2003 ���չ��� ���·� ������ ����ϼ���.');</script>"
    end if
    response.write "<script>history.back();</script>"
    response.end
end if

response.write "OK"
response.end

''������ ó��.
okCNT = 0

dbget.BeginTrans
    for i=0 to UBound(xlRowALL)

    if IsObject(xlRowALL(i)) then
        set iLine = xlRowALL(i)
        ''�ɼǰ��� �ִ°��.

		if IsNumeric(iLine.FItemArray(1)) then
			'// 20120101 or 120101
			if (Len(CStr(iLine.FItemArray(1))) = 8) then
				iLine.FItemArray(1) = Left(iLine.FItemArray(1), 4) & "-" & Mid(iLine.FItemArray(1), 5, 2) & "-" & Right(iLine.FItemArray(1), 2)
			elseif (Len(CStr(iLine.FItemArray(1))) = 6) then
				iLine.FItemArray(1) = Left(iLine.FItemArray(1), 2) & "-" & Mid(iLine.FItemArray(1), 3, 2) & "-" & Right(iLine.FItemArray(1), 2)
			end if
		end if

		if IsNumeric(iLine.FItemArray(2)) then
			'// 20120101 or 120101
			if (Len(CStr(iLine.FItemArray(2))) = 8) then
				iLine.FItemArray(2) = Left(iLine.FItemArray(2), 4) & "-" & Mid(iLine.FItemArray(2), 5, 2) & "-" & Right(iLine.FItemArray(2), 2)
			elseif (Len(CStr(iLine.FItemArray(2))) = 6) then
				iLine.FItemArray(2) = Left(iLine.FItemArray(2), 2) & "-" & Mid(iLine.FItemArray(2), 3, 2) & "-" & Right(iLine.FItemArray(2), 2)
			end if
		end if

        iLine.FItemArray(17) = Replace(iLine.FItemArray(17),",","")
        iLine.FItemArray(18) = Replace(iLine.FItemArray(18),",","")
        iLine.FItemArray(19) = Replace(iLine.FItemArray(19),",","")

        IF (UBound(iLine.FItemArray)>30) and (iLine.FItemArray(18)<>"") then
            IF (iLine.FItemArray(30)="") then iLine.FItemArray(30)="0"
            IF (iLine.FItemArray(31)="") then iLine.FItemArray(31)="0"

            iLine.FItemArray(18) = CLNG(iLine.FItemArray(18)) + CLNG(iLine.FItemArray(30))
        END IF

        iLine.FItemArray(3) = convPayTypeStr2Code(iLine.FItemArray(3))
        IF (iLine.FItemArray(3)="") then iLine.FItemArray(3)="50"                           ''PayType
        IF (iLine.FItemArray(2)="") then iLine.FItemArray(2)=iLine.FItemArray(1)            ''Paydate
        IF (iLine.FItemArray(19)="") then iLine.FItemArray(19)=iLine.FItemArray(18)         ''RealSellPrice

        '''�����ȣ�� �ּ�1�� ������� [ ]�� �����ȣ ���� = �ϳ�������.
        iLine.FItemArray(12) = TRIM(Replace(iLine.FItemArray(12),"  "," "))
        iLine.FItemArray(13) = TRIM(Replace(iLine.FItemArray(13),"  "," "))
        iLine.FItemArray(14) = TRIM(Replace(iLine.FItemArray(14),"  "," "))
        IF (iLine.FItemArray(12)=iLine.FItemArray(13)) then
            POS1 = InStr(iLine.FItemArray(12),"[")
            POS2 = InStr(iLine.FItemArray(12),"]")
            IF (POS1>0) and (POS2>0) then
                iLine.FItemArray(12) = Mid(iLine.FItemArray(12),POS1+1,POS2-POS1-1)
                iLine.FItemArray(12) = Trim (iLine.FItemArray(12))

                IF (iLine.FItemArray(13)=iLine.FItemArray(14)) THEN
                    iLine.FItemArray(13) = TRIM(Mid(iLine.FItemArray(13),POS2+1,512))
                    iLine.FItemArray(14) = iLine.FItemArray(13)
                ELSE
                    iLine.FItemArray(13) = TRIM(Mid(iLine.FItemArray(13),POS2+1,512))
                END IF
            END IF
        END IF

        '''�ּҿ� ���ּҰ� ������� 3��° Blank���� ����.
        POS1 = 0
        POS2 = 0
        POS3 = 0

        IF (iLine.FItemArray(13)=iLine.FItemArray(14)) then
            POS1 = InStr(iLine.FItemArray(14)," ")
            ''rw "POS1="&POS1
            IF (POS1>0) then
                POS2 = InStr(MID(iLine.FItemArray(14),POS1+1,512)," ")
                ''rw "POS2="&POS2
                IF POS2>0 then
                    POS3 = InStr(MID(iLine.FItemArray(14),POS1+POS2+1,512)," ")
                    IF POS3>0 then
                        iLine.FItemArray(13)=LEFT(iLine.FItemArray(14),POS1+POS2+POS3-1)
                        iLine.FItemArray(14)=MID(iLine.FItemArray(14),POS1+POS2+POS3+1,512)

                        'rw iLine.FItemArray(13)
                        'rw iLine.FItemArray(14)
                    END IF
                END IF
            END IF
        END IF

		'�ֹ���ȣ-����(17709513-1)
		''if (SellSite = 9) then
		''	if (Ubound(Split(iLine.FItemArray(0), "-")) > 0) then
		''		iLine.FItemArray(0) = Split(iLine.FItemArray(0), "-")(0)
		''	end if
		''end if

		if False then
			rw "@tplcompanyid="&tplcompanyid
			rw "@sellsitename="&iLine.FItemArray(32)
			rw "@OutMallOrderSerial="&iLine.FItemArray(0)
			rw "@SellDate="&iLine.FItemArray(1)
			rw "@PayType="&iLine.FItemArray(3)
			rw "@Paydate="&iLine.FItemArray(2)
			rw "@partnerItemID="&iLine.FItemArray(15)
			rw "@partnerItemName="&iLine.FItemArray(21)
			rw "@partnerOption="&iLine.FItemArray(16)
			rw "@partnerOptionName="&iLine.FItemArray(22)
			rw "@OrderUserID="&iLine.FItemArray(4)
			rw "@OrderName="&iLine.FItemArray(5)
			rw "@OrderEmail="&iLine.FItemArray(8)
			rw "@OrderTelNo="&iLine.FItemArray(6)
			rw "@OrderHpNo="&iLine.FItemArray(7)
			rw "@ReceiveName="&iLine.FItemArray(9)
			rw "@ReceiveTelNo="&iLine.FItemArray(10)
			rw "@ReceiveHpNo="&iLine.FItemArray(11)
			rw "@ReceiveZipCode="&iLine.FItemArray(12)
			rw "@ReceiveAddr1="&iLine.FItemArray(13)
			rw "@ReceiveAddr2="&iLine.FItemArray(14)
			rw "@SellPrice="&iLine.FItemArray(18)
			rw "@RealSellPrice="&iLine.FItemArray(19)
			rw "@ItemOrderCount="&iLine.FItemArray(17)
			rw "@OrgDetailKey="&iLine.FItemArray(25)
			rw "@deliverymemo="&iLine.FItemArray(26)
			rw "@requireDetail="&iLine.FItemArray(27)
			rw "barcode=" & iLine.FItemArray(31)
		end if

		IF (iLine.FItemArray(0)<>"") then
			partnerItemName = iLine.FItemArray(21)
			partnerOptionName = iLine.FItemArray(22)
			barcode = iLine.FItemArray(31)
			''partnerItemName = iLine.FItemArray(33)
			''partnerOptionName = iLine.FItemArray(34)


			itemgubun = ""
			itemid = ""
			itemoption = ""
			itemname = ""
			itemoptionname = ""

			if barcode <> "" then
				'// �����ڵ� �ִ� ���
				'// 1. �����ڵ�� ��ǰ�ڵ� ã��
				sqlStr = " select top 1 "
				sqlStr = sqlStr & " 	b.itemgubun, b.itemid, b.itemoption "
				sqlStr = sqlStr & " 	, IsNull(i.itemname, si.shopitemname) as itemname, IsNull(IsNull(o.optionname, si.shopitemoptionname), '') as itemoptionname "
				''sqlStr = sqlStr & " 	, IsNull(i.orgprice, si.orgsellprice) as orgprice "
				''sqlStr = sqlStr & " 	, IsNull(i.mainimage, si.offimgmain) as mainimage, IsNull(i.listimage, si.offimglist) as listimage, IsNull(i.smallimage, si.offimgsmall) as smallimage "
				sqlStr = sqlStr & " from "
				sqlStr = sqlStr & " 	[db_item].[dbo].[tbl_item_option_stock] b "
				sqlStr = sqlStr & " 	left join [db_item].[dbo].[tbl_item] i "
				sqlStr = sqlStr & " 	on "
				sqlStr = sqlStr & " 		1 = 1 "
				sqlStr = sqlStr & " 		and b.itemgubun = '10' "
				sqlStr = sqlStr & " 		and b.itemid = i.itemid "
				sqlStr = sqlStr & " 	left join [db_item].[dbo].[tbl_item_option] o "
				sqlStr = sqlStr & " 	on "
				sqlStr = sqlStr & " 		i.itemid = o.itemid and b.itemoption = o.itemoption "
				sqlStr = sqlStr & " 	left join [db_shop].[dbo].[tbl_shop_item] si "
				sqlStr = sqlStr & " 	on "
				sqlStr = sqlStr & " 		1 = 1 "
				sqlStr = sqlStr & " 		and b.itemgubun <> '10' "
				sqlStr = sqlStr & " 		and b.itemgubun = si.itemgubun "
				sqlStr = sqlStr & " 		and b.itemid = si.shopitemid "
				sqlStr = sqlStr & " 		and b.itemoption = si.itemoption "
				sqlStr = sqlStr & " 	left join [db_partner].[dbo].[tbl_partner] ip on i.makerid = ip.id "
				sqlStr = sqlStr & " 	left join [db_partner].[dbo].[tbl_partner] sp on si.makerid = sp.id "
				sqlStr = sqlStr & " where b.barcode = '" & barcode & "' and IsNull(ip.id, sp.id) = '" & makerid & "' "
				''rw sqlStr
				rsTENget.CursorLocation = adUseClient
				rsTENget.Open sqlStr,dbTENget,adOpenForwardOnly, adLockReadOnly
				if  not rsTENget.EOF  then
					itemgubun         = rsTENget("itemgubun")
					itemid            = rsTENget("itemid")
					itemoption        = rsTENget("itemoption")
					itemname          = rsTENget("itemname")
					itemoptionname    = rsTENget("itemoptionname")
					''orgprice          = rsTENget("orgprice")
					''mainimage         = rsTENget("mainimage")
					''listimage         = rsTENget("listimage")
					''smallimage        = rsTENget("smallimage")
				end if
				rsTENget.Close
			end if

			if (itemgubun <> "") then
				''rw "Found : " & barcode
			else
				''rw "Not Found" & partnerItemName
			end if

			'// ================================================================
			'// ���� : ������ ������ ���ν����� �����ؾ� �Ѵ�.!!!
			'// ================================================================
            paramInfo = Array(Array("@RETURN_VALUE",adInteger,adParamReturnValue,,0) _
                ,Array("@tplcompanyid" , adVarchar	, adParamInput,32, tplcompanyid)	_
                ,Array("@SellSite" , adInteger	, adParamInput,, "")	_
				,Array("@SellSiteName" , adVarchar	, adParamInput,32, iLine.FItemArray(32))	_
				,Array("@OutMallOrderSerial"	, adVarchar	, adParamInput,23, Left(CStr(iLine.FItemArray(0)), 22))	_
				,Array("@OrgDetailKey"	,adVarchar, adParamInput,32, iLine.FItemArray(25)) _
    			,Array("@SellDate"	,adDate, adParamInput,, iLine.FItemArray(1)) _
    			,Array("@PayType"	,adVarchar, adParamInput,32, iLine.FItemArray(3)) _
    			,Array("@Paydate"	,adDate, adParamInput,, iLine.FItemArray(2)) _
				,Array("@makerid" , adVarchar	, adParamInput,32, makerid)	_
				,Array("@itemgubun" , adVarchar	, adParamInput,2, itemgubun)	_
				,Array("@itemid" , adInteger	, adParamInput,, itemid)	_
				,Array("@itemoption" , adVarchar	, adParamInput,4, itemoption)	_
				,Array("@itemname" , adVarchar	, adParamInput,128, itemname)	_
				,Array("@itemoptionname" , adVarchar	, adParamInput,128, itemoptionname)	_
    			,Array("@orderItemID"	,adVarchar, adParamInput,32, iLine.FItemArray(15)) _
    			,Array("@orderItemName"	,adVarchar, adParamInput,128, iLine.FItemArray(21)) _
    			,Array("@orderItemOption"	,adVarchar, adParamInput,128, iLine.FItemArray(16)) _
    			,Array("@orderItemOptionName"	,adVarchar, adParamInput,128, iLine.FItemArray(22)) _
    			,Array("@barcode"	,adVarchar, adParamInput,32, barcode) _
    			,Array("@OrderName"	,adVarchar, adParamInput,32, iLine.FItemArray(5)) _
    			,Array("@OrderEmail"	,adVarchar, adParamInput,100, iLine.FItemArray(8)) _
    			,Array("@OrderTelNo"	,adVarchar, adParamInput,16, iLine.FItemArray(6)) _
    			,Array("@OrderHpNo"	,adVarchar, adParamInput,16, iLine.FItemArray(7)) _
    			,Array("@ReceiveName"	,adVarchar, adParamInput,32, iLine.FItemArray(9)) _
    			,Array("@ReceiveTelNo"	,adVarchar, adParamInput,16, iLine.FItemArray(10)) _
    			,Array("@ReceiveHpNo"	,adVarchar, adParamInput,16, iLine.FItemArray(11)) _
    			,Array("@ReceiveZipCode"	,adVarchar, adParamInput,7, iLine.FItemArray(12)) _
    			,Array("@ReceiveAddr1"	,adVarchar, adParamInput,128, iLine.FItemArray(13)) _
    			,Array("@ReceiveAddr2"	,adVarchar, adParamInput,512, iLine.FItemArray(14)) _
    			,Array("@SellPrice"	,adCurrency, adParamInput,, iLine.FItemArray(18)) _
    			,Array("@RealSellPrice"	,adCurrency, adParamInput,, iLine.FItemArray(19)) _
				,Array("@vatinclude"	,adVarchar, adParamInput,1, "Y") _
    			,Array("@ItemOrderCount"	,adInteger, adParamInput,, iLine.FItemArray(17)) _
    			,Array("@DeliveryType"	,adInteger, adParamInput,, 0) _
    			,Array("@deliveryprice"	,adCurrency, adParamInput,, 0) _
    			,Array("@deliverymemo"	,adVarchar, adParamInput,400, iLine.FItemArray(26)) _
				,Array("@countryCode"	,adVarchar, adParamInput,2, "KR") _
    			,Array("@requireDetail"	,adVarchar, adParamInput,400, iLine.FItemArray(27)) _
    			,Array("@retErrStr"	,adVarchar, adParamOutput,100, "") _
    			)

''''             0�ֹ���ȣ, �ֹ���, �Ա���, ���Ҽ���, �ֹ���ID, �ֹ���, �ֹ�����ȭ,�ֹ����޴���ȭ,�ֹ����̸���,
''''             9������,��������ȭ,�������ڵ���,������Zip,������addr1,������addr2
''''             15��ǰ�ڵ�, �ɼ��ڵ�, ����, �ǸŰ�, �Һ��ڰ�, �����, ��ǰ��, �ɼǸ�,
''''             23��ü��ǰ�ڵ�,��ü�ɼ��ڵ�,�ֹ�������Ű,�ֹ����ǻ���, ��ǰ�䱸����1, ��ǰ�䱸����2, ��ǰ�䱸����3

            sqlStr = "db_threepl.dbo.usp_OnlineTmpOrder_Insert"
            retParamInfo = fnExecSPOutput(sqlStr,paramInfo)

            RetErr    = GetValue(retParamInfo, "@RETURN_VALUE") ' �����ڵ�
            retErrStr  = GetValue(retParamInfo, "@retErrStr") ' ������ �����ȣ

            okCNT = okCNT +1

        END IF
            set iLine = Nothing

            IF (retErr)<>0 then
                dbget.rollbackTrans
                response.write "ERROR["&retErr&"]"& retErrStr
                response.write "<script>alert('"&Replace("ERROR["&retErr&"]"& retErrStr,"'","")&"');</script>"
                response.write "<script>history.back();</script>"
                response.end
            END IF
        end if
    Next
dbget.CommitTrans

response.write "<script>alert('"&okCNT&"�� �ԷµǾ����ϴ�.')</script>"
response.write "<script>opener.location.reload();self.close();</script>"
'''====================================================================================

Class TXLRowObj
    public FItemArray

    public function setArrayLength(ln)
        Redim FItemArray(ln)
    end function

End Class

function convPayTypeStr2Code(oStr)
    SELECT CASE oStr
        CASE "�ſ�ī��" : convPayTypeStr2Code="100"
        CASE "�ſ�" : convPayTypeStr2Code="100"
        CASE "������" : convPayTypeStr2Code="7"
        CASE "�ǽð���ü" : convPayTypeStr2Code="20"
        CASE "�ڵ�������" : convPayTypeStr2Code="400"
        CASE "�޴�������" : convPayTypeStr2Code="400"
        CASE "�ڵ���" : convPayTypeStr2Code="400"
        CASE "�޴���" : convPayTypeStr2Code="400"
        CASE ELSE : convPayTypeStr2Code="50"
    END SELECT
end function

function IsSKipRow(ixlRow, skipCol0Str)

    if Not IsArray(ixlRow) then
        IsSKipRow = true
        Exit function
    end if

    if  LCASE(ixlRow(0))=LCASE(skipCol0Str) then
        IsSKipRow = true
        Exit function
    end if

    IsSKipRow = false
end function

Function fnGetXLFileArray(byref xlRowALL, sFilePath, aSheetName, iArrayLen)
    Dim conDB, Rs, strQry, iResult, i, J, iObj
    Dim irowObj, strTable
    '' on Error ���� ���� �ȵ�.. ���� ���ѷ��� ���µ�.

    Set conDB = Server.CreateObject("ADODB.Connection")
    conDB.Provider = "Microsoft.Jet.oledb.4.0"
    conDB.Properties("ExtEnded Properties").Value = "Excel 8.0;"

    On Error Resume Next
        conDB.Open sFilePath

        IF (ERR) then
            fnGetXLFileArray=false
            exit function
        End if
    On Error Goto 0

    '' get First Sheet Name=============''��Ʈ�� �������ΰ�� ������ �� ����.
    Set Rs = conDB.OpenSchema(adSchemaTables)

    IF Not Rs.Eof Then
        aSheetName = Rs.Fields("table_name").Value
        ''rw "aSheetName="&aSheetName
    ENd IF
    Set Rs = Nothing
    ''==================================

    Set Rs = Server.CreateObject("ADODB.Recordset")

    ''strQry = "Select * From [sheet1$]"
    strQry = "Select * From ["&aSheetName&"]"

    ReDim xlRowALL(0)
    fnGetXLFileArray = true

On Error Resume Next
    Rs.Open strQry, conDB
    	IF (ERR) then
            fnGetXLFileArray=false
            Rs.Close
            Set Rs = Nothing
            Set conDB = Nothing
            exit function
        End if

        If Not Rs.Eof Then
            Do Until Rs.Eof
                IF (ERR) then
                    fnGetXLFileArray=false
                    Rs.Close
                    Set Rs = Nothing
                    Set conDB = Nothing
                    exit function
                End if

                set irowObj = new TXLRowObj
                irowObj.setArrayLength(iArrayLen)

                For i=0 to ArrayLen
                    if (xlPosArr(i)<0) then
                        irowObj.FItemArray(i) = ""
                    else
                        irowObj.FItemArray(i) = Replace(null2blank(Rs(xlPosArr(i))),"*","")
                    end if
                Next

                IF (Not IsSKipRow(irowObj.FItemArray,skipString)) then
                    ReDim Preserve xlRowALL(UBound(xlRowALL)+1)

                    set xlRowALL(UBound(xlRowALL)) =  irowObj
                    ''xlRowALL(UBound(xlRowALL)).arrayObj = xlRow

                END IF
                set irowObj = Nothing
                Rs.MoveNext
            Loop
       else
          fnGetXLFileArray=false
       end if

       ''''On Error Goto 0
        IF (ERR) then
            fnGetXLFileArray=false
        End if
    Rs.Close
On Error Goto 0

    Set Rs = Nothing
    Set conDB = Nothing

    if Ubound(xlRowALL)< 1 then fnGetXLFileArray=false

End Function

Function AddTmpDbOrderData(ixlRowALL)
    AddTmpDbOrderData = false

'    Rs.Open strQry, conDB
'        If Not Rs.Eof Or Rs.Bof Then
'            Do Until Rs.Eof
'            		iResult =""
'            		'Add1 =""
'            		'Add2=""
'            		outmallorderserial 	= replace(Rs(1),"*","")	'�ֹ���ȣ
'            		partnerItemId 		= Rs(3)	'��ǰ�ڵ�
'            		partnerItemName 	= ReplaceRequestSpecialChar(Rs(4))	'��ǰ��
'            		partnerOptionName	= replace(replace(Rs(5),"-","")," ","")	'�ɼ�
'            		SellPrice			= Rs(6)	'�ֹ��ݾ�
'            		ItemOrderCount 	= Rs(7)	'�ֹ�����
'            		OrderName		= Rs(9)	'�̸�
'            		ZipCode			= left(Rs(11),3)&right(Rs(11),3)	'�����ȣ
'            		arrAdd			= split(Rs(12)," ")	'�������ּ�
'            		for i = 0 to 2
'            		Add1			= Add1 &" "& arrAdd(i)  	'�������ּ�1
'            		next
'            		for i = 3 to ubound(arrAdd)
'            		Add2			= Add2 &" "& arrAdd(i)  	'�������ּ�1
'            		next
'            		IF Add2 = "" THEN Add2 = "."
'            		ReceiveName		= Rs(13)	'������
'            		ReceiveTelNo		= Rs(14)	'��ȭ��ȣ
'            		ReceiveHpNo		= Rs(15)	'�ڵ���
'            		EtcAsk			= ReplaceRequestSpecialChar(Rs(16))	'�ֹ���û����
'            		PayDate			= Rs(18)	'����(�Ա�)����
'            		RealSellPrice		= Rs(21)	'�Ǹűݾ�
'
'            		SellSite			="3"
'            		PartnerSeq		="58"
'            		OrderEmail 		=""
'
''iResult =  clsConnDB.fnMultiExecSPReturnValue("db_agirlOrder.dbo.[usp_Back_OutMallOrder_Insert]("&SellSite&","&PartnerSeq &" ,'"&OutMallOrderSerial&"','"&partnerItemID&"','"&partnerItemName&"','','"&partnerOptionName&"','"&OrderName&"'"&_
''    				",'"&OrderEmail&"','"&ReceiveTelNo&"','"&ReceiveHpNo&"','"&ReceiveName&"','"&ReceiveTelNo&"','"&ReceiveHpNo&"','"&ZipCode&"','"&Add1&"','"&Add2&"','"&EtcAsk&"','"&SellPrice&"','"&RealSellPrice&"','"&PayDate&"',1,'"&ItemOrderCount&"')")
'
'    		IF iResult  = 0 THEN
%>

<%
''clsConnDB.RollbackTrans
''Set clsConnDB = nothing
'			Set Rs = Nothing
'			Set conDB = Nothing
'		ELSEIF iResult = 2 THEN
'			if  errMsg1 ="" then
'			errMsg1 = OutMallOrderSerial
'			else
'			errMsg1 = errMsg1 &","&OutMallOrderSerial
'			end if
'		ELSEIF iResult = 3 THEN
'			if  errMsg2 ="" then
'			errMsg2 = partnerItemID
'			else
'			errMsg2 = errMsg2 &","&partnerItemID
'			end if
'    		END IF
'            Rs.MoveNext
'            Loop
'         end if
            'clsConnDB.CommitTrans
            'Set clsConnDB = nothing
'    Set Rs = Nothing
'    Set conDB = Nothing
end Function
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db_TPLClose.asp" -->
