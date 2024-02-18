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
    ''뉴한국택배물류 01 
''대한통운 02 
''삼성택배 03 
''로엑스(구 아주) 04 
''옐로우캡택배 05 
''우체국택배 06 
''우편등기 07 
''이클라인 08 
''주코택배 09 
''트라넷택배 10 
''한진택배 11 
''현대택배 12 
''동부(구훼미리) 13 
''CJ택배 14 
''KGB 15 
''후다닥(퀵) 16 
''우리택배 17 
''건영택배 18 
''제니엘택배 19 
''코덱스택배 20 
''호남택배 21 
''일양택배 22 
''로젠택배 24 
''스마일택배 25 
''대신화물택배 26 
''동원로엑스택배 27 
''천일택배  28 
''꼬모택배 29 
''경동택배 30 
''벨익스프레스 31 
''리브로배송 32 
''고려택배 33 
''wpx택배 34 
''쎄덱스택배 35 
''고려택배 36 
''사가와익스프레스 37 
''하나로택배 38 
''정안물류 39 
''EMS국제우편 40 
''굿모닝택배 41 
''용마로지스 42 
''이노지스택배 43 
''네덱스 44 
function TenDlvCode2DnshopDlvCode(itenCode)
    select Case itenCode
        CASE 1 : TenDlvCode2DnshopDlvCode = "11"     ''한진
        CASE 2 : TenDlvCode2DnshopDlvCode = "12"     ''현대
        CASE 3 : TenDlvCode2DnshopDlvCode = "02"     ''대한통운
        CASE 4 : TenDlvCode2DnshopDlvCode = "14"     ''CJ GLS
        CASE 5 : TenDlvCode2DnshopDlvCode = "08"     ''이클라인
        CASE 6 : TenDlvCode2DnshopDlvCode = "03"     ''삼성 HTH
        CASE 7 : TenDlvCode2DnshopDlvCode = "13"     ''동부(구훼미리)
        CASE 8 : TenDlvCode2DnshopDlvCode = "06"     ''우체국택배
        CASE 9 : TenDlvCode2DnshopDlvCode = "15"     ''KGB택배
        CASE 10 : TenDlvCode2DnshopDlvCode = "04"     ''아주택배 / 로엑스(구 아주)
        CASE 11 : TenDlvCode2DnshopDlvCode = ""     ''오렌지택배
        CASE 12 : TenDlvCode2DnshopDlvCode = "01"     ''한국택배 / 뉴한국택배물류?
        CASE 13 : TenDlvCode2DnshopDlvCode = "05"     ''옐로우캡
        CASE 14 : TenDlvCode2DnshopDlvCode = ""     ''나이스택배
        CASE 15 : TenDlvCode2DnshopDlvCode = ""     ''중앙택배
        CASE 16 : TenDlvCode2DnshopDlvCode = "09"     ''주코택배
        CASE 17 : TenDlvCode2DnshopDlvCode = "10"     ''트라넷택배
        CASE 18 : TenDlvCode2DnshopDlvCode = "24"     ''로젠택배
        CASE 19 : TenDlvCode2DnshopDlvCode = "15"     ''KGB특급택배
        CASE 20 : TenDlvCode2DnshopDlvCode = "20"     ''KT로지스
        CASE 21 : TenDlvCode2DnshopDlvCode = "30"     ''경동택배
        CASE 22 : TenDlvCode2DnshopDlvCode = "33"     ''고려택배
        CASE 23 : TenDlvCode2DnshopDlvCode = "35"     ''쎄덱스택배 신세계
        CASE 24 : TenDlvCode2DnshopDlvCode = "37"     ''사가와익스프레스
        CASE 25 : TenDlvCode2DnshopDlvCode = "38"     ''하나로택배
        CASE 26 : TenDlvCode2DnshopDlvCode = "22"     ''일양택배
        CASE 27 : TenDlvCode2DnshopDlvCode = "04"     ''LOEX택배
        CASE 28 : TenDlvCode2DnshopDlvCode = "13"     ''동부익스프레스
        CASE 29 : TenDlvCode2DnshopDlvCode = "18"     ''건영택배
        CASE 30 : TenDlvCode2DnshopDlvCode = "43"     ''이노지스
        CASE 31 : TenDlvCode2DnshopDlvCode = "28"     ''천일택배
        CASE 33 : TenDlvCode2DnshopDlvCode = "21"     ''호남택배
        CASE 99 : TenDlvCode2DnshopDlvCode = "00"     ''업체직송
        CASE  Else
            TenDlvCode2DnshopDlvCode = ""
    end Select
end function
    
function TenDlvCode2InterParkDlvCode(itenCode)
    select Case itenCode
        CASE 1 : TenDlvCode2InterParkDlvCode = "169178"     ''한진
        CASE 2 : TenDlvCode2InterParkDlvCode = "169198"     ''현대
        CASE 3 : TenDlvCode2InterParkDlvCode = "169177"     ''대한통운
        CASE 4 : TenDlvCode2InterParkDlvCode = "169168"     ''CJ GLS
        CASE 5 : TenDlvCode2InterParkDlvCode = "169211"     ''이클라인
        CASE 6 : TenDlvCode2InterParkDlvCode = "169181"     ''삼성 HTH
        CASE 7 : TenDlvCode2InterParkDlvCode = ""     ''동부(구훼미리)
        CASE 8 : TenDlvCode2InterParkDlvCode = "169199"     ''우체국택배
        CASE 9 : TenDlvCode2InterParkDlvCode = "169187"     ''KGB택배
        CASE 10 : TenDlvCode2InterParkDlvCode = "169194"     ''아주택배 / 로엑스(구 아주)
        CASE 11 : TenDlvCode2InterParkDlvCode = ""     ''오렌지택배
        CASE 12 : TenDlvCode2InterParkDlvCode = ""     ''한국택배 / 뉴한국택배물류?
        CASE 13 : TenDlvCode2InterParkDlvCode = "169200"     ''옐로우캡
        CASE 14 : TenDlvCode2InterParkDlvCode = ""     ''나이스택배
        CASE 15 : TenDlvCode2InterParkDlvCode = ""     ''중앙택배
        CASE 16 : TenDlvCode2InterParkDlvCode = ""     ''주코택배
        CASE 17 : TenDlvCode2InterParkDlvCode = ""     ''트라넷택배
        CASE 18 : TenDlvCode2InterParkDlvCode = "169182"     ''로젠택배
        CASE 19 : TenDlvCode2InterParkDlvCode = ""     ''KGB특급택배
        CASE 20 : TenDlvCode2InterParkDlvCode = ""     ''KT로지스
        CASE 21 : TenDlvCode2InterParkDlvCode = "303978"     ''경동택배
        CASE 22 : TenDlvCode2InterParkDlvCode = "169526"     ''고려택배
        CASE 23 : TenDlvCode2InterParkDlvCode = "236288"     ''쎄덱스택배 신세계
        CASE 24 : TenDlvCode2InterParkDlvCode = "231491"     ''사가와익스프레스
        CASE 25 : TenDlvCode2InterParkDlvCode = "229381"     ''하나로택배
        CASE 26 : TenDlvCode2InterParkDlvCode = "263792"     ''일양택배
        CASE 27 : TenDlvCode2InterParkDlvCode = "169194"     ''LOEX택배
        CASE 28 : TenDlvCode2InterParkDlvCode = "231145"     ''동부익스프레스
        CASE 29 : TenDlvCode2InterParkDlvCode = "231194"     ''건영택배
        CASE 30 : TenDlvCode2InterParkDlvCode = "266237"     ''이노지스
        CASE 31 : TenDlvCode2InterParkDlvCode = "230175"     ''천일택배
        CASE 33 : TenDlvCode2InterParkDlvCode = "250701"     ''호남택배
        CASE 99 : TenDlvCode2InterParkDlvCode = "169167"     ''업체직송/기타
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
	
	''MakedSongjangStr = "주문번호,협력사번호,상점번호,택배사번호,송장번호,주문인," + VbCrlf + MakedSongjangStr
	
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

                    '상품명 매칭(및 이름이 다른것 변환)
                    if instr(tmpItemNm,"연아의 다이어리 꿈을 꾸다 2010")>0 then tmpItemNm = "연아의 다이어리 <꿈을 꾸다>"
               		tmpItemNm = Trim(Replace(tmpItemNm,"[텐바이텐]",""))
                    
                    ''' 중간에 브랜드명이 들어감.. 브랜드 삭제..
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
    
    
    ''MakedSongjangStr = 순번	주문번호	주문일련번호	주문자	수령인	상품명	상품옵션	입금확인일	택배업체코드	발송량	송장번호
    MakedSongjangStr = VBCRLF+ VBCRLF+ "순번,주문번호,주문일련번호,주문자,수령인,상품명,상품옵션,입금확인일,택배업체코드,발송량,송장번호" + VbCrlf + MakedSongjangStr
    
    response.write MakedSongjangStr
end if

%>

<!-- #include virtual="/lib/db/dbclose.asp" -->