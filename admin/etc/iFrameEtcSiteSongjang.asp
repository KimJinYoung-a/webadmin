<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->

<%
'Response.Buffer=False
Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=" & request("sitename") & FormatDateTime(now(),2) & "_" & Replace(FormatDateTime(now(),4),":","") & ".csv"
Response.CacheControl = "public"

function ReplaceSCVStr(oStr)
    ReplaceSCVStr = ""
    if IsNULL(oStr) then Exit function
    ReplaceSCVStr = Replace(oStr, chr(34),"'")
    
end function

Class CExtSiteSongJangItem
    ''���ѱ��ù蹰�� 01 
''������� 02 
''�Ｚ�ù� 03 
''�ο���(�� ����) 04 
''���ο�ĸ�ù� 05 
''��ü���ù� 06 
''������ 07 
''��Ŭ���� 08 
''�����ù� 09 
''Ʈ����ù� 10 
''�����ù� 11 
''�����ù� 12 
''����(���ѹ̸�) 13 
''CJ�ù� 14 
''KGB 15 
''�Ĵٴ�(��) 16 
''�츮�ù� 17 
''�ǿ��ù� 18 
''���Ͽ��ù� 19 
''�ڵ����ù� 20 
''ȣ���ù� 21 
''�Ͼ��ù� 22 
''�����ù� 24 
''�������ù� 25 
''���ȭ���ù� 26 
''�����ο����ù� 27 
''õ���ù�  28 
''�����ù� 29 
''�浿�ù� 30 
''���ͽ������� 31 
''����ι�� 32 
''����ù� 33 
''wpx�ù� 34 
''�굦���ù� 35 
''����ù� 36 
''�簡���ͽ������� 37 
''�ϳ����ù� 38 
''���ȹ��� 39 
''EMS�������� 40 
''�¸���ù� 41 
''�븶������ 42 
''�̳������ù� 43 
''�׵��� 44 
function TenDlvCode2DnshopDlvCode(itenCode)
    select Case itenCode
        CASE 1 : TenDlvCode2DnshopDlvCode = "11"     ''����
        CASE 2 : TenDlvCode2DnshopDlvCode = "12"     ''����
        CASE 3 : TenDlvCode2DnshopDlvCode = "02"     ''�������
        CASE 4 : TenDlvCode2DnshopDlvCode = "14"     ''CJ GLS
        CASE 5 : TenDlvCode2DnshopDlvCode = "08"     ''��Ŭ����
        CASE 6 : TenDlvCode2DnshopDlvCode = "03"     ''�Ｚ HTH
        CASE 7 : TenDlvCode2DnshopDlvCode = "13"     ''����(���ѹ̸�)
        CASE 8 : TenDlvCode2DnshopDlvCode = "06"     ''��ü���ù�
        CASE 9 : TenDlvCode2DnshopDlvCode = "15"     ''KGB�ù�
        CASE 10 : TenDlvCode2DnshopDlvCode = "04"     ''�����ù� / �ο���(�� ����)
        CASE 11 : TenDlvCode2DnshopDlvCode = ""     ''�������ù�
        CASE 12 : TenDlvCode2DnshopDlvCode = "01"     ''�ѱ��ù� / ���ѱ��ù蹰��?
        CASE 13 : TenDlvCode2DnshopDlvCode = "05"     ''���ο�ĸ
        CASE 14 : TenDlvCode2DnshopDlvCode = ""     ''���̽��ù�
        CASE 15 : TenDlvCode2DnshopDlvCode = ""     ''�߾��ù�
        CASE 16 : TenDlvCode2DnshopDlvCode = "09"     ''�����ù�
        CASE 17 : TenDlvCode2DnshopDlvCode = "10"     ''Ʈ����ù�
        CASE 18 : TenDlvCode2DnshopDlvCode = "24"     ''�����ù�
        CASE 19 : TenDlvCode2DnshopDlvCode = "15"     ''KGBƯ���ù�
        CASE 20 : TenDlvCode2DnshopDlvCode = "20"     ''KT������
        CASE 21 : TenDlvCode2DnshopDlvCode = "30"     ''�浿�ù�
        CASE 22 : TenDlvCode2DnshopDlvCode = "33"     ''����ù�
        CASE 23 : TenDlvCode2DnshopDlvCode = "35"     ''�굦���ù� �ż���
        CASE 24 : TenDlvCode2DnshopDlvCode = "37"     ''�簡���ͽ�������
        CASE 25 : TenDlvCode2DnshopDlvCode = "38"     ''�ϳ����ù�
        CASE 26 : TenDlvCode2DnshopDlvCode = "22"     ''�Ͼ��ù�
        CASE 27 : TenDlvCode2DnshopDlvCode = "04"     ''LOEX�ù�
        CASE 28 : TenDlvCode2DnshopDlvCode = "13"     ''�����ͽ�������
        CASE 29 : TenDlvCode2DnshopDlvCode = "18"     ''�ǿ��ù�
        CASE 30 : TenDlvCode2DnshopDlvCode = "43"     ''�̳�����
        CASE 31 : TenDlvCode2DnshopDlvCode = "28"     ''õ���ù�
        CASE 33 : TenDlvCode2DnshopDlvCode = "21"     ''ȣ���ù�
        CASE 99 : TenDlvCode2DnshopDlvCode = "00"     ''��ü����
        CASE  Else
            TenDlvCode2DnshopDlvCode = ""
    end Select
end function
    
function TenDlvCode2InterParkDlvCode(itenCode)
    select Case itenCode
        CASE 1 : TenDlvCode2InterParkDlvCode = "169178"     ''����
        CASE 2 : TenDlvCode2InterParkDlvCode = "169198"     ''����
        CASE 3 : TenDlvCode2InterParkDlvCode = "169177"     ''�������
        CASE 4 : TenDlvCode2InterParkDlvCode = "169168"     ''CJ GLS
        CASE 5 : TenDlvCode2InterParkDlvCode = "169211"     ''��Ŭ����
        CASE 6 : TenDlvCode2InterParkDlvCode = "169181"     ''�Ｚ HTH
        CASE 7 : TenDlvCode2InterParkDlvCode = ""     ''����(���ѹ̸�)
        CASE 8 : TenDlvCode2InterParkDlvCode = "169199"     ''��ü���ù�
        CASE 9 : TenDlvCode2InterParkDlvCode = "169187"     ''KGB�ù�
        CASE 10 : TenDlvCode2InterParkDlvCode = "169194"     ''�����ù� / �ο���(�� ����)
        CASE 11 : TenDlvCode2InterParkDlvCode = ""     ''�������ù�
        CASE 12 : TenDlvCode2InterParkDlvCode = ""     ''�ѱ��ù� / ���ѱ��ù蹰��?
        CASE 13 : TenDlvCode2InterParkDlvCode = "169200"     ''���ο�ĸ
        CASE 14 : TenDlvCode2InterParkDlvCode = ""     ''���̽��ù�
        CASE 15 : TenDlvCode2InterParkDlvCode = ""     ''�߾��ù�
        CASE 16 : TenDlvCode2InterParkDlvCode = ""     ''�����ù�
        CASE 17 : TenDlvCode2InterParkDlvCode = ""     ''Ʈ����ù�
        CASE 18 : TenDlvCode2InterParkDlvCode = "169182"     ''�����ù�
        CASE 19 : TenDlvCode2InterParkDlvCode = ""     ''KGBƯ���ù�
        CASE 20 : TenDlvCode2InterParkDlvCode = ""     ''KT������
        CASE 21 : TenDlvCode2InterParkDlvCode = "303978"     ''�浿�ù�
        CASE 22 : TenDlvCode2InterParkDlvCode = "169526"     ''����ù�
        CASE 23 : TenDlvCode2InterParkDlvCode = "236288"     ''�굦���ù� �ż���
        CASE 24 : TenDlvCode2InterParkDlvCode = "231491"     ''�簡���ͽ�������
        CASE 25 : TenDlvCode2InterParkDlvCode = "229381"     ''�ϳ����ù�
        CASE 26 : TenDlvCode2InterParkDlvCode = "263792"     ''�Ͼ��ù�
        CASE 27 : TenDlvCode2InterParkDlvCode = "169194"     ''LOEX�ù�
        CASE 28 : TenDlvCode2InterParkDlvCode = "231145"     ''�����ͽ�������
        CASE 29 : TenDlvCode2InterParkDlvCode = "231194"     ''�ǿ��ù�
        CASE 30 : TenDlvCode2InterParkDlvCode = "266237"     ''�̳�����
        CASE 31 : TenDlvCode2InterParkDlvCode = "230175"     ''õ���ù�
        CASE 33 : TenDlvCode2InterParkDlvCode = "250701"     ''ȣ���ù�
        CASE 99 : TenDlvCode2InterParkDlvCode = "169167"     ''��ü����/��Ÿ
        CASE  Else
            TenDlvCode2InterParkDlvCode = ""
    end Select
end function

    ''authcode+ ',' + 'B540'+ ',' + '1'+ ',' + '37'+ ',' +  deliverno+ ',' + buyname + ','
    
    public FExtOrderNo
    public FShopCode
    public FShopSeq
    public Fbuyname
    public Freqname
    public FSongjangDiv
    public Fdeliverno
    public FDlvCNT
    
    public FOrgSeq
    public FDetailSeq
    public FItemName
    public FItemOptionName
    public FIpkumdate
    
    public function GetInterParkSongJangStr()
        dim extSongjangDiv 
        GetInterParkSongJangStr = ""
        extSongjangDiv = TenDlvCode2InterParkDlvCode(FSongjangDiv)
        if (extSongjangDiv<>"") and (Not IsNULL(Fdeliverno)) and (Fdeliverno<>"") then
            GetInterParkSongJangStr = FOrgSeq + "," + FExtOrderNo + "," + FDetailSeq + "," + Chr(34) + FBuyName + Chr(34) + "," + Chr(34) + FReqName + Chr(34) + "," + Chr(34) + FItemName + Chr(34) + "," + Chr(34) + FItemOptionName + Chr(34) + "," + FIpkumdate + "," + extSongjangDiv +  "," + CStr(FDlvCNT) + "," + Fdeliverno + ""
        end if
    end function

    public function GetDnShopSongJangStr()
        dim extSongjangDiv 
        GetDnShopSongJangStr = ""
        extSongjangDiv = TenDlvCode2DnshopDlvCode(FSongjangDiv)
        if (extSongjangDiv<>"") and (Not IsNULL(Fdeliverno)) and (Fdeliverno<>"") then
            GetDnShopSongJangStr = FExtOrderNo + "," + FShopCode + "," + FShopSeq + "," + extSongjangDiv + "," + Replace(Fdeliverno,"-","") + "," + Fbuyname + ","
        end if
    end function

    Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

    End Sub

end Class

dim sitename, orgFile
sitename = requestCheckVar(request("sitename"),32)
orgFile  = request("orgFile")

dim iLines, iBufStr, iBufStr2
dim i,j,cnt
dim iExtOrderList, iExtOrderNo
dim sqlStr

dim FResultCount, FItemList()
dim StRegDate
dim pmaxCnt, pSongjangStr, MakedSongjangStr
dim DnshopShopSeq
dim tmpItemNm
DnshopShopSeq = ""

Dim Pos1

if (sitename="dnshop") then
    iBufStr = Split(orgFile,VbCrlf)
    
    if ISArray(iBufStr) then
        for i=LBound(iBufStr) to UBound(iBufStr)
            iLines = iBufStr(i)
            
            if (Trim(Left(iLines,15))<>"") then
                iExtOrderList = iExtOrderList + "'" + Trim(Left(iLines,15)) + "'" + ","
                
                if (DnshopShopSeq="") then
                    DnshopShopSeq = TRim(split(iLines,VbTab)(2))
                end if
            end if
        next
    end if
    
    if Right(iExtOrderList,1)="," then iExtOrderList = Left(iExtOrderList,Len(iExtOrderList)-1)
    
    StRegDate = Left(CStr(DateAdd("m",-1,Now())),10)
    
    sqlStr = " select m.orderserial, m.deliverno, m.authcode, m.buyname, d.songjangdiv, d.songjangno, count(d.idx) as CNT"
	sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m,"
	sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d"
	sqlStr = sqlStr + " where m.orderserial=d.orderserial"
	'sqlStr = sqlStr + " and m.regdate>'" + StRegDate + "'"
	sqlStr = sqlStr + " and m.sitename='dnshop'"
	sqlStr = sqlStr + " and m.cancelyn='N'"
	sqlStr = sqlStr + " and d.cancelyn='N'"
	sqlStr = sqlStr + " and d.itemid<>0"
	sqlStr = sqlStr + " and m.ipkumdiv>7"
	sqlStr = sqlStr + " and m.authcode in ("
	sqlStr = sqlStr + " " + iExtOrderList 
	sqlStr = sqlStr + " )"
	sqlStr = sqlStr + " group by m.orderserial, m.authcode, m.deliverno, m.buyname, d.songjangdiv, d.songjangno"
	sqlStr = sqlStr + " order by  m.orderserial, m.deliverno desc"

	rsget.Open sqlStr, dbget, 1
	
    FResultCount = rsget.RecordCount
    redim preserve FItemList(FResultCount)

    if  not rsget.EOF  then
		j = 0
		do until rsget.eof
			set FItemList(j) = new CExtSiteSongJangItem
			
			FItemList(j).FExtOrderNo   = rsget("authcode")
            FItemList(j).FShopCode     = "B540" 
            FItemList(j).FShopSeq      = DnshopShopSeq  '' "1"
            FItemList(j).Fbuyname      = db2Html(rsget("buyname"))
            FItemList(j).FSongjangDiv  = rsget("songjangdiv")
            FItemList(j).Fdeliverno    = rsget("songjangno")
            FItemList(j).FDlvCNT       = rsget("CNT")
			rsget.MoveNext
			j = j + 1
		loop
	end if
	
	rsget.close
	
	if ISArray(iBufStr) then
        for i=LBound(iBufStr) to UBound(iBufStr)
            iLines = iBufStr(i)
            pmaxCnt = 0
            pSongjangStr = ""
            if (Trim(Left(iLines,15))<>"") then
                iExtOrderNo = Trim(Left(iLines,15))
                
                for j=0 to FResultCount-1
                    if (iExtOrderNo=FItemList(j).FExtOrderNo) then
                        
                        if (FItemList(j).FDlvCNT>pmaxCnt) then
                            pSongjangStr = FItemList(j).GetDnShopSongJangStr
                            pmaxCnt = FItemList(j).FDlvCNT
                        end if
                    end if
                next
                
                if (pSongjangStr<>"") then
                    MakedSongjangStr = MakedSongjangStr + pSongjangStr + VbCrlf
                end if
            end if
        next
    end if
	
	''MakedSongjangStr = "�ֹ���ȣ,���»��ȣ,������ȣ,�ù���ȣ,�����ȣ,�ֹ���," + VbCrlf + MakedSongjangStr
	
	response.write MakedSongjangStr
elseif (sitename="interpark") then
    iBufStr = Split(orgFile,VbCrlf)
    
    if ISArray(iBufStr) then
        for i=LBound(iBufStr) to UBound(iBufStr)
            iLines = iBufStr(i)
            iBufStr2 = Split(iLines,VbTab)
            if ISArray(iBufStr2) then
                if (Trim(Left(iLines,15))<>"") then
                    iExtOrderList = iExtOrderList + "'" + Trim(iBufStr2(1)) + "'" + ","
                end if
            end if
        next
    end if
    
    
    if Right(iExtOrderList,1)="," then iExtOrderList = Left(iExtOrderList,Len(iExtOrderList)-1)
    
    StRegDate = Left(CStr(DateAdd("m",-1,Now())),10)
    
    sqlStr = " select m.orderserial, m.authcode, d.itemname, d.songjangdiv, d.songjangno, d.currstate, d.itemno"
	sqlStr = sqlStr + " from [db_order].[dbo].tbl_order_master m,"
	sqlStr = sqlStr + " [db_order].[dbo].tbl_order_detail d"
	sqlStr = sqlStr + " where m.orderserial=d.orderserial"
	'sqlStr = sqlStr + " and m.regdate>'" + StRegDate + "'"
	sqlStr = sqlStr + " and m.sitename='interpark'"
	sqlStr = sqlStr + " and m.cancelyn='N'"
	sqlStr = sqlStr + " and d.cancelyn='N'"
	sqlStr = sqlStr + " and d.itemid<>0"
	sqlStr = sqlStr + " and m.ipkumdiv>4"
	sqlStr = sqlStr + " and m.authcode in ("
	sqlStr = sqlStr + " " + iExtOrderList 
	sqlStr = sqlStr + " )"
	sqlStr = sqlStr + " order by  m.orderserial desc"

	rsget.Open sqlStr, dbget, 1
	
    FResultCount = rsget.RecordCount
    redim preserve FItemList(FResultCount)

    if  not rsget.EOF  then
		j = 0
		do until rsget.eof
			set FItemList(j) = new CExtSiteSongJangItem
			
			FItemList(j).FExtOrderNo   = rsget("authcode")
            FItemList(j).FSongjangDiv  = rsget("songjangdiv")
            FItemList(j).Fdeliverno    = rsget("songjangno")
            FItemList(j).FDlvCNT       = rsget("itemno")
            FItemList(j).FItemName     = db2html(rsget("itemname"))
			rsget.MoveNext
			j = j + 1
		loop
	end if
	
	rsget.close
	
	if ISArray(iBufStr) then
        for i=LBound(iBufStr) to UBound(iBufStr)
            iLines = iBufStr(i)
            pmaxCnt = 0
            pSongjangStr = ""
            
            iBufStr2 = Split(iLines,VbTab)
            
            if IsArray(iBufStr2) then
                if (UBound(iBufStr2)>5) then
                    iExtOrderNo = Trim(iBufStr2(1))
                    tmpItemNm = Trim(iBufStr2(5))

                    '��ǰ�� ��Ī(�� �̸��� �ٸ��� ��ȯ)
                    if instr(tmpItemNm,"������ ���̾ ���� �ٴ� 2010")>0 then tmpItemNm = "������ ���̾ <���� �ٴ�>"
               		tmpItemNm = Trim(Replace(tmpItemNm,"[�ٹ�����]",""))
                    
                    ''' �߰��� �귣����� ��.. �귣�� ����..
                    Pos1 = InStr(tmpItemNm," ")
                    IF (Pos1>0) then
                        tmpItemNm = TRIM(Mid(tmpItemNm,Pos1,255))
                    End IF
                    
                    for j=0 to FResultCount-1
                        if (iExtOrderNo=FItemList(j).FExtOrderNo) then
                            
                            if (FItemList(j).FItemName=tmpItemNm) then
                                FItemList(j).FOrgSeq =  Trim(iBufStr2(0))
                                FItemList(j).FDetailSeq   =  Trim(iBufStr2(2))
                                FItemList(j).FItemName   =  Trim(iBufStr2(5))
                                FItemList(j).FItemOptionName   =  Trim(iBufStr2(6))
                                
                                FItemList(j).FBuyName	=  Trim(iBufStr2(3))
                                FItemList(j).FReqName	=  Trim(iBufStr2(4))
                                FItemList(j).FIpkumdate  =  Trim(iBufStr2(7))
                                IF Trim(iBufStr2(9))<>"" then
	                                FItemList(j).FDlvCNT =  Trim(iBufStr2(9))
	                            end IF
                            
                                pSongjangStr = FItemList(j).GetInterParkSongJangStr
                            end if
                        end if
                    next
                    
                    if (pSongjangStr<>"") then
                        MakedSongjangStr = MakedSongjangStr + pSongjangStr + VbCrlf
                    end if
                end if
            end if
        next
    end if
    
    
    ''MakedSongjangStr = ����	�ֹ���ȣ	�ֹ��Ϸù�ȣ	�ֹ���	������	��ǰ��	��ǰ�ɼ�	�Ա�Ȯ����	�ù��ü�ڵ�	�߼۷�	�����ȣ
    MakedSongjangStr = VBCRLF+ VBCRLF+ "����,�ֹ���ȣ,�ֹ��Ϸù�ȣ,�ֹ���,������,��ǰ��,��ǰ�ɼ�,�Ա�Ȯ����,�ù��ü�ڵ�,�߼۷�,�����ȣ" + VbCrlf + MakedSongjangStr
    
    response.write MakedSongjangStr
end if

%>

<!-- #include virtual="/lib/db/dbclose.asp" -->