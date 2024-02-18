<%
public function getTrackNaverURIByTrName(isongjangdiv,isongjangno)
    dim itrname : itrname = fnreplaceNvTrName(getSongjangDiv2Val(isongjangdiv,1) )
    if isNULL(itrname) or isNULL(isongjangno) then
        getTrackNaverURIByTrName = "#"
    else
        getTrackNaverURIByTrName = "https://search.naver.com/search.naver?query=" + itrname + "+" + TRIM(replace(isongjangno,"-",""))
    end if
end function

public function fnreplaceNvTrName(trname)
    dim retnm
    retnm = trname
    if isNULL(retnm) then
        retnm = ""
    else
        retnm = replace(retnm,"(구)","")
        retnm = replace(retnm,"CVSnet택배","CVSnet")
        retnm = replace(retnm,"CU Post","CU 편의점택배")
        retnm = replace(retnm,"대신화물택배","대신택배")
    end if
    fnreplaceNvTrName = retnm
end Function

public function getBrandAvgDeliverInfo(iStartDate,iEndDate,iMakerid,iEtcdivinc)
    Dim sqlStr
    sqlStr = "exec [db_dataSummary].[dbo].[usp_TEN_delveryTrack_Summary_BrandInfo]  '"&iStartDate&"','"&iEndDate&"','"&iMakerid&"',"&iEtcdivinc&""
    
    db3_dbget.CursorLocation = adUseClient
    db3_rsget.Open sqlStr,db3_dbget,adOpenForwardOnly,adLockReadOnly
    If not db3_rsget.EOF Then
        getBrandAvgDeliverInfo = db3_rsget.getRows()
    End If
    db3_rsget.Close

end function

function getBrandDefaultDlv(imakerid)
    Dim strSql
    strSql = "select defaultsongjangdiv from db_partner.dbo.tbl_partner WITH(NOLOCK) where id='"&imakerid&"'"
    dbget.CursorLocation = adUseClient
    rsget.Open strSql,dbget,adOpenForwardOnly,adLockReadOnly
	If not rsget.EOF Then
		getBrandDefaultDlv = rsget("defaultsongjangdiv")
	End If
	rsget.Close

end function

function getSongjangDlvBoxHtml(idivcode,iboxname,iaddstr)
    Dim strSql, ArrRows, i
    Dim ret

    if isNULL(idivcode) then idivcode="" ''널인경우

    ArrRows = session("songjangDivArr")
    if NOT isArray(ArrRows) then
        strSql = " exec [db_dataSummary].[dbo].[usp_TEN_sonngjangdiv_GET] "
        db3_dbget.CursorLocation = adUseClient
        db3_rsget.Open strSql,db3_dbget,adOpenForwardOnly,adLockReadOnly
        If not db3_rsget.EOF Then
            ArrRows = db3_rsget.getRows()
            session("songjangDivArr") = ArrRows
        End If
        db3_rsget.Close
    end if

    ret = "<select name='"&iboxname&"' "&iaddstr&">"
    ret = ret&"<option value='' "&CHKIIF(idivcode="","selected","")&">선택</option>"
    if isArray(ArrRows) then
        For i=0 To UBound(ArrRows,2)
            if (CStr(idivcode)=CStr(ArrRows(0,i))) then
                ret = ret&"<option value='"&ArrRows(0,i)&"' selected >"&ArrRows(1,i)&"</option>"
            else
                if (ArrRows(3,i)="Y") then ''사용택배사만 뿌리자.
                    ret = ret&"<option value='"&ArrRows(0,i)&"'>"&ArrRows(1,i)&"</option>"
                end if
            end if
		Next	
    end if 
    ret = ret&"</select>"

    getSongjangDlvBoxHtml = ret
end function

function getSongjangDiv2Val(idivcode,ivalColum)
    Dim strSql, ArrRows, i

    ''ivalColum 1:Name, 2:URL
    if isNULL(idivcode) then Exit function
    getSongjangDiv2Val = idivcode
    if NOT (ivalColum=1 or ivalColum=2) then Exit function

    ArrRows = session("songjangDivArr")
    if NOT isArray(ArrRows) then
        strSql = " exec [db_dataSummary].[dbo].[usp_TEN_sonngjangdiv_GET] "
        db3_dbget.CursorLocation = adUseClient
        db3_rsget.Open strSql,db3_dbget,adOpenForwardOnly,adLockReadOnly
        If not db3_rsget.EOF Then
            ArrRows = db3_rsget.getRows()
            session("songjangDivArr") = ArrRows
        End If
        db3_rsget.Close
    end if
    
    if isArray(ArrRows) then
        For i=0 To UBound(ArrRows,2)
            if (CStr(ArrRows(0,i))=CStr(idivcode)) then
                getSongjangDiv2Val = ArrRows(ivalColum,i)
                Exit function
            end if
		Next	
    end if
end function

Sub drawTrackDeliverBox(iboxname,iselname, iptype)
    dim ret, strSql, arrVal, i
    strSql = " exec [db_dataSummary].[dbo].[usp_TEN_sonngjangdiv_GET] '"&iptype&"'"
        
    db3_dbget.CursorLocation = adUseClient
    db3_rsget.Open strSql,db3_dbget,adOpenForwardOnly,adLockReadOnly
	If not db3_rsget.EOF Then
		arrVal = db3_rsget.getRows()
	End If
	db3_rsget.Close
	
	if isArray(arrVal) then
	    ret = "<select name='"&iboxname&"' id='"&iboxname&"'>"
	    ret = ret&"<option value='' "&CHKIIF(iselname="","selected","")&">선택</option>"
	    for i=0 To UBound(arrVal,2)
	        ret = ret&"<option value='"&arrVal(0,i)&"' "&CHKIIF(iselname=CStr(arrVal(0,i)),"selected","")&">"&arrVal(1,i)&"</option>"
	    next
	    ret = ret&"</select>"
	end if
	response.write ret
end Sub

Class CCSDeliveryTrackBrandSumItem

    public Fmakerid   
    public FTTLDLVFIN
    public FdelayTTL
    public FmibeaTTL
    public FmijiphaTTL
    public FjiphaNoMoveTTL

    public FdelayTTLOrderGrp
    public FmibeaTTLOrderGrp
    public FmijiphaTTLOrderGrp
    public FjiphaNoMoveTTLOrderGrp
    

    Private Sub Class_Initialize()
		'//
    End Sub
    Private Sub Class_Terminate()
		'//
    End Sub
end Class

Class CCSDeliveryTrackDetailItem
    public Forderserial
    public Fsongjangno
    public Fsongjangdiv
    public Fmakerid
    public Fcurrstate
    public Fipkumdiv
    public Fjumundiv
    public Fipkumdate
    public Fupcheconfirmdate
    public Fbeasongdate
    public Fdeparturedt
    public FdlvfinishDT
    public Fjungsanfixdate
    public Ftraceupddt
    public Flastupdate
    public Fisupchebeasong
    public Fdivname
    public FsongjangTrackURL

    public Ftracetbltype
    public FregDt
    public FtraceAcctCnt
    public Fchktype

    public FqueRegdt
    public Fquelastupddt
    public Fquelastupdno

    public Ftrdeparturedt
    public Ftrarrivedt
    public Ftrupddt

    public Fodetailidx              
    public Fbuyname
    public Freqname
    public Freqzipaddr
    public FSitename
    public Fitemid
    public Fitemname
    public Fitemoptionname

    public Freguserid

    public FcsCNT 
    public FcsFinCNT 

    public function getErrChkTypeName()
        getErrChkTypeName=""
        if isNULL(Fchktype) then exit function
        SELECT CASE Fchktype
            'CASE 1
            '    : getErrChkTypeName = "재추적"
            'CASE 2
            '    : getErrChkTypeName = "배송일"
            CASE 4
                : getErrChkTypeName = "택배사"
            CASE 5
                : getErrChkTypeName = "DIGIT"
            CASE 9
                : getErrChkTypeName = "길이"
            CASE else :
                getErrChkTypeName = CStr(Fchktype)
        END SELECT
    end function

    '' asp mod in 2,147,483,647
    Private Function fnMod2(Value1, Value2)
        fnMod2 = Value1 - (FIX(Value1 / Value2) * Value2)
    End Function

    public Function getOrderDtlStatusName()
        dim ret : ret=""
        if (Fjumundiv="6") then
            ret = "교환<br>"
        elseif  (Fjumundiv="9") then
            ret = "반품<br>"
        else

        end if

        SELECT CASE Fcurrstate
            CASE 0 :
                ret = ret&""
            CASE 3 :
                ret = ret&"업체확인"
            CASE 7 :
                ret = ret&"출고완료"
                
            CASE else
                ret = ret& Fcurrstate
        END SELECT
        
        getOrderDtlStatusName = ret
    end function 

    public Function fnExtractNumber(ByVal iInputString)
        Dim i 
        Dim retNum 
        For i = Len(iInputString) To 1 Step -1
            If IsNumeric(Mid(iInputString, i, 1)) Then
                retNum = Mid(iInputString, i, 1) & retNum
            End If
        Next 
        fnExtractNumber = retNum
    End Function


    public function getDefaultLengthOfSongjangDiv(isongjangdiv)
        getDefaultLengthOfSongjangDiv = -1
        if isNULL(isongjangdiv) then Exit function

        Select CASE isongjangdiv
            CASE "3","4" '' cj대한통운
                : getDefaultLengthOfSongjangDiv = 12
            CASE "1","2" '' 한진, 롯데.
                : getDefaultLengthOfSongjangDiv = 12
            CASE "18" '' 로젠
                : getDefaultLengthOfSongjangDiv = 11
            CASE "8" '' 우체국
                : getDefaultLengthOfSongjangDiv = 13
            CASE "21" '' 경동택배
                : getDefaultLengthOfSongjangDiv = 13
            CASE else
                : getDefaultLengthOfSongjangDiv = -1
        End Select
    end function

    public function getSecondLengthOfSongjangDiv(isongjangdiv)
        getSecondLengthOfSongjangDiv = -1
        if isNULL(isongjangdiv) then Exit function

        Select CASE isongjangdiv
            CASE "3","4" '' cj대한통운
                : getSecondLengthOfSongjangDiv = 10
            CASE else
                : getSecondLengthOfSongjangDiv = -1
        End Select
    end function

    public function getMayDigitCode(isongjangdiv,isongjangno)
        dim buf
        if isNULL(isongjangdiv) then Exit function
        if isNULL(isongjangno) then Exit function

        
        if Len(isongjangno)<1 then Exit function

        '' 대한통운
        if (CStr(isongjangdiv)="3") or (CStr(isongjangdiv)="4") then
            ''if NOT (Len(isongjangno)=12 or Len(isongjangno)<10) then Exit function
            if NOT (Len(isongjangno)=12) then Exit function
            ''digit 제외한 뒤 9자리를 7로 나눈 나머지.
            buf = CCUR(LEFT(RIGHT(isongjangno,10),9))
            getMayDigitCode = CStr(fnMod2(buf,7))
        elseif (CStr(isongjangdiv)="1") or (CStr(isongjangdiv)="2") then
            if NOT (Len(isongjangno)=12) then Exit function
            ''digit 왼쪽 11 자리를 7로 나눈 나머지
            buf = CCUR(LEFT(isongjangno,11)) 
            getMayDigitCode = CStr(fnMod2(buf,7))
        elseif (CStr(isongjangdiv)="18") then
            if NOT (Len(isongjangno)=11) then Exit function
            ''digit 왼쪽 10 자리를 7로 나눈 나머지
            buf = CCUR(LEFT(isongjangno,10)) 
            getMayDigitCode = CStr(fnMod2(buf,7))
        else
            getMayDigitCode = ""
        end if 
    end function

    public function isMaySongjangDivBySongjangno(isongjangdiv,isongjangno)
        isMaySongjangDivBySongjangno = false
        if isNULL(isongjangdiv) then Exit function
        if isNULL(isongjangno) then Exit function

        if (getDefaultLengthOfSongjangDiv(isongjangdiv)<>LEN(isongjangno)) and (getSecondLengthOfSongjangDiv(isongjangdiv)<>LEN(isongjangno)) then Exit function

        dim left2Digit : left2Digit=LEFT(isongjangno,2)
        '' 대한통운
        if (CStr(isongjangdiv)="3") or (CStr(isongjangdiv)="4") then
            isMaySongjangDivBySongjangno = (left2Digit="34") or (left2Digit="62") or (left2Digit="35") or (left2Digit="33") or (left2Digit="38") or (left2Digit="36") or (left2Digit="65")
        elseif (CStr(isongjangdiv)="1") then
            isMaySongjangDivBySongjangno = (left2Digit="41") or (left2Digit="50")
        elseif (CStr(isongjangdiv)="2") then
            isMaySongjangDivBySongjangno = (left2Digit="23") or (left2Digit="40")
        elseif (CStr(isongjangdiv)="18") then
            isMaySongjangDivBySongjangno = (left2Digit="90") or (left2Digit="93") or (left2Digit="94") or (left2Digit="95") or (left2Digit="96")
        elseif (CStr(isongjangdiv)="8") then
            isMaySongjangDivBySongjangno = (left2Digit="68") or (left2Digit="60") 
        else
            isMaySongjangDivBySongjangno = false
        end if 

    end function

    public function isMayOtherDliver(byref imayOtherDeliverCode,isongjangdiv,isongjangno)
        if isNULL(isongjangdiv) then Exit function
        if isNULL(isongjangno) then Exit function

        isMayOtherDliver = FALSE

        '' 길이가 맞고 Digit 코드가 맞으면 return
        if ((getDefaultLengthOfSongjangDiv(isongjangdiv)=LEN(isongjangno) or getSecondLengthOfSongjangDiv(isongjangdiv)=LEN(isongjangno)) and getMayDigitCode(isongjangdiv,isongjangno)=RIGHT(isongjangno,1) ) then Exit function

        if (isongjangdiv="3" or isongjangdiv="4") then
            if ( isMaySongjangDivBySongjangno(18,isongjangno) and getMayDigitCode(18,isongjangno)=RIGHT(isongjangno,1) ) then
                imayOtherDeliverCode = 18
                isMayOtherDliver = true
                Exit function
            end if

            if ( isMaySongjangDivBySongjangno(1,isongjangno) and getMayDigitCode(1,isongjangno)=RIGHT(isongjangno,1) ) then
                imayOtherDeliverCode = 1
                isMayOtherDliver = true
                Exit function
            end if

            if ( isMaySongjangDivBySongjangno(2,isongjangno) and getMayDigitCode(2,isongjangno)=RIGHT(isongjangno,1) ) then
                imayOtherDeliverCode = 2
                isMayOtherDliver = true
                Exit function
            end if

            
        elseif (isongjangdiv="1" or isongjangdiv="2") then
            if ( isMaySongjangDivBySongjangno(18,isongjangno) and getMayDigitCode(18,isongjangno)=RIGHT(isongjangno,1) ) then
                imayOtherDeliverCode = 18
                isMayOtherDliver = true
                Exit function
            end if

            if ( isMaySongjangDivBySongjangno(4,isongjangno) and getMayDigitCode(4,isongjangno)=RIGHT(isongjangno,1) ) then
                imayOtherDeliverCode = 4
                isMayOtherDliver = true
                Exit function
            end if
        elseif (isongjangdiv="18") then
            if ( isMaySongjangDivBySongjangno(4,isongjangno) and getMayDigitCode(4,isongjangno)=RIGHT(isongjangno,1) ) then
                imayOtherDeliverCode = 4
                isMayOtherDliver = true
                Exit function
            end if

            if ( isMaySongjangDivBySongjangno(2,isongjangno) and getMayDigitCode(2,isongjangno)=RIGHT(isongjangno,1) ) then
                imayOtherDeliverCode = 2
                isMayOtherDliver = true
                Exit function
            end if

            if ( isMaySongjangDivBySongjangno(1,isongjangno) and getMayDigitCode(1,isongjangno)=RIGHT(isongjangno,1) ) then
                imayOtherDeliverCode = 1
                isMayOtherDliver = true
                Exit function
            end if
        elseif (isongjangdiv="8") then
            if ( isMaySongjangDivBySongjangno(3,isongjangno) and getMayDigitCode(4,isongjangno)=RIGHT(isongjangno,1) ) then
                imayOtherDeliverCode = 4
                isMayOtherDliver = true
                Exit function
            end if

            if ( isMaySongjangDivBySongjangno(2,isongjangno) and getMayDigitCode(2,isongjangno)=RIGHT(isongjangno,1) ) then
                imayOtherDeliverCode = 2
                isMayOtherDliver = true
                Exit function
            end if

            if ( isMaySongjangDivBySongjangno(1,isongjangno) and getMayDigitCode(1,isongjangno)=RIGHT(isongjangno,1) ) then
                imayOtherDeliverCode = 1
                isMayOtherDliver = true
                Exit function
            end if
        end if
    end function

    public function getTrackStateUpcheView()
        Dim retStr : retStr = ""

        if Not isNULL(Ftrdeparturedt) then  'oDeliveryTrackFake.FItemList(i).
            retStr = "<strong>집하완료</strong>"
        end if

        if Not isNULL(Ftrarrivedt) then  'oDeliveryTrackFake.FItemList(i).
            retStr = "<strong>배송완료</strong><br>("&Ftrarrivedt&")"  'oDeliveryTrackFake.FItemList(i).
        end if

        if (Fquelastupdno>=200) then
            if (retStr="") then
                retStr = "추적불가"
            end if
            retStr = retStr & "<br>재시도회수("&Fquelastupdno&")<br>최종추적:"&LEFT(Fquelastupddt,10)
        end if 

        if (Fquelastupdno<200) and (retStr="") then
            retStr = "추적중<br>재시도회수("&Fquelastupdno&")"
            if (Fquelastupdno<5) then
                retStr="추적중"
            end if
        end if 

        getTrackStateUpcheView = retStr
    end function

    public function getDigitChkStr
        if isNULL(Fsongjangno) then Exit function
        if isNULL(Fsongjangdiv) then Exit function

        dim isongjangno : isongjangno = fnExtractNumber(Fsongjangno)
        dim maylen : maylen = getDefaultLengthOfSongjangDiv(Fsongjangdiv)
        dim maylen2 : maylen2 = getSecondLengthOfSongjangDiv(Fsongjangdiv)
        dim maydigit : maydigit = getMayDigitCode(Fsongjangdiv,isongjangno)
        dim currDigit : currDigit = Right(isongjangno,1)
        dim imayOtherDeliverCode
        dim retStr : retStr =""

        if (maylen<>LEN(isongjangno) and maylen2<>LEN(isongjangno)) then
            if (maylen>0) then
                retStr = retStr & "길이오류:("&LEN(isongjangno)&"/"&maylen&")"
            end if
        end if

        if (maydigit<>currDigit) and (maydigit<>"") then
            if (retStr<>"") then retStr=retStr&"<br>"
            retStr = retStr & "번호검증오류" ''&maydigit
        end if

        if isMayOtherDliver(imayOtherDeliverCode,Fsongjangdiv,isongjangno) then
            if (retStr<>"") then retStr=retStr&"<br>"
            retStr = retStr & ""&getSongjangDiv2Val(imayOtherDeliverCode,1)&"?"
        end if

        getDigitChkStr = retStr
    end function

    public function getDlvDivName()
        getDlvDivName = Fdivname
        if isNULL(Fdivname) or (Fdivname="") then
            getDlvDivName = Fsongjangdiv
        end if
    end function

    public function getDlvDivName2()
        getDlvDivName2 = Fdivname
        if isNULL(Fdivname) or (Fdivname="") then
            getDlvDivName2 = getSongjangDiv2Val(Fsongjangdiv,1)
        end if
    end function

    public function getTraceTBLTypeName()
        if isNULL(Ftracetbltype) then Exit function

        if Ftracetbltype=1 then
            getTraceTBLTypeName = "추적결과"
        elseif Ftracetbltype=2 then
            getTraceTBLTypeName = "추적Que"
        elseif Ftracetbltype=3 then
            getTraceTBLTypeName = "추적QueErr"
        end if
    end function


    public function isValidPopTraceSongjangDiv()
        isValidPopTraceSongjangDiv = false

        if isNULL(FsongjangTrackURL) then Exit function
        if (Trim(FsongjangTrackURL)="") then Exit function

        isValidPopTraceSongjangDiv = true
    end function

    public function getTrackURI()
        if isNULL(FsongjangTrackURL) or isNULL(Fsongjangno)  then
            getTrackURI = "#"
        else
            getTrackURI = FsongjangTrackURL + TRIM(replace(Fsongjangno,"-",""))
        end if
    end function

    public function getTrackNaverURI()
        dim idivname
        if isNULL(Fdivname) or isNULL(Fsongjangno)  then
            getTrackNaverURI = "#"
        else
            idivname = Fdivname
            if Fsongjangdiv="3" then idivname="CJ대한통운" ''(구)대한통운
            if Fsongjangdiv="35" then idivname="CVSnet" ''CVSnet택배,CU Post 
            if Fsongjangdiv="42" then idivname="CU 편의점택배" ''CU Post
            if Fsongjangdiv="34" then idivname="대신택배" ''대신화물택배

            getTrackNaverURI = "https://search.naver.com/search.naver?query=" + idivname + "+" + TRIM(replace(Fsongjangno,"-",""))
        end if
    end function



    public function getUpbeaGubunName()
        if (Fisupchebeasong="Y") then
            getUpbeaGubunName = "업배"
        elseif (Fisupchebeasong="N") then
            getUpbeaGubunName = "<strong>텐배</strong>"
        else
            getUpbeaGubunName = Fisupchebeasong
        end if
    end function
    
    Private Sub Class_Initialize()
		'//
        FcsCNT = 0
        FcsFinCNT = 0
    End Sub
    Private Sub Class_Terminate()
		'//
    End Sub
end Class

Class CDeliveryTrackSmrItem
    '' beasongdate	songjangdiv	makerid	isupchebeasong	ttlchulgono	jiphafinCNT	dlvfinCNT	D-N일	D+0일	D+1일	D+2일	D+3일	미집하	미배송	divname
    public Fbeasongdate
    public Fsongjangdiv
    public Fmakerid
    public Fisupchebeasong
    public Fttlchulgono
    public FjiphafinCNT
    public FdlvfinCNT

    public FDminusCnt
    public FDplus0Cnt
    public FDplus1Cnt
    public FDplus2Cnt
    public FDplus3UpCnt

    public FMijiphaCnt
    public FMidlvfinCnt

    public FsongjangDivName
    public FsongjangTrackURL
    public FErrChkCnt

    public function getMijipHaPro()
        if (Fttlchulgono<>0) then
            getMijipHaPro = FIX((1-FMijiphaCnt*1.0/Fttlchulgono)*100) & " %"
        end if
    end function

    public function getMiBeasongPro()
        if (Fttlchulgono<>0) then
            getMiBeasongPro = FIX((1-FMidlvfinCnt*1.0/Fttlchulgono)*100) & " %"
        end if
    end function

    public function getUpbeaGubunName()
        if (Fisupchebeasong="Y") then
            getUpbeaGubunName = "업배"
        elseif (Fisupchebeasong="N") then
            getUpbeaGubunName = "<strong>텐배</strong>"
        else
            getUpbeaGubunName = Fisupchebeasong
        end if
    end function

    public function getSongjangDivName()
        if (isNULL(FsongjangDivName)) then
            getSongjangDivName = Fsongjangdiv
            if (getSongjangDivName="-1") then getSongjangDivName=""
        else
            getSongjangDivName = FsongjangDivName
        end if
    end function

    Private Sub Class_Initialize()
		'//
    End Sub
    Private Sub Class_Terminate()
		'//
    End Sub
End Class

Class CDeliverySongjangChangeLogItem
    public Fsongjangchgidx
    public Fodetailidx
    public Forderserial
    public Fpsongjangno
    public Fpsongjangdiv
    public Fchgsongjangno
    public Fchgsongjangdiv
    public Fchguserid
    public Fregdt
    public Fupddt

    public FactionType
    public Fchgdlvfinishdt
    public Fitemid
    public Fitemoption
    public Fsongjangno
    public Fsongjangdiv
    public Fbeasongdate
    public Fdlvfinishdt
    public Fjungsanfixdate

    public Fitemno
    public Fmakerid
    public Fitemname
    public Fitemoptionname
    
    public Fdcancelyn
    public Fsitename
    public Fcancelyn
    public Fauthcode
    public Freqzipaddr
    public Fcomment




    Private Sub Class_Initialize()
		'//
    End Sub
    Private Sub Class_Terminate()
		'//
    End Sub
End Class


Class CDeliveryTrack
    public FItemList()
    public FOneItem
    public FCurrPage
    public FTotalPage
    public FPageSize
    public FResultCount
    public FScrollCount
    public FTotalCount

    public FRectStartDate
	public FRectEndDate
	public FRectSongjangDiv
	public FRectOrderSerial
    public FRectMakerid
    public FRectIsUpchebeasong
    public FRectSongjangNo

    public FRectGrpBeasongdate
    public FRectGrpSongjangDiv
    public FRectGrpBrand

    public FRectMijipHaExists
    public FRectMiBeasongExists
    public FRectEtcdivinc
    public FRectErrChkType 
    public FRectByList

    public FSumdelayTTL
    public FSumMibeaTTL
    public FSummijiphaTTL
    public FSumjiphaNoMoveTTL

    public FSumdelayTTLOrderGrp
    public FSummibeaTTLOrderGrp
    public FSummijiphaTTLOrderGrp
    public FSumjiphaNoMoveTTLOrderGrp

    public FRectSitename
    public FRectSiteScope
    public FRectNotIncComment
    public FRectNotIncMapXjungsan

    public FRectSearchType

    public function getSongjangChangeLogListWithCmt()
        Dim sqlStr
        sqlStr = "exec [db_dataSummary].[dbo].[usp_Ten_Check_SongjangChangeLog] '"&FRectStartDate&"','"&FRectEndDate&"','"&FRectSitename&"',"&CHKIIF(FRectSiteScope="","NULL",FRectSiteScope)&","&CHKIIF(FRectNotIncComment="","NULL","1")&","&CHKIIF(FRectNotIncMapXjungsan="","NULL","1")&""

        db3_dbget.CursorLocation = adUseClient
        db3_rsget.Open sqlStr,db3_dbget,adOpenForwardOnly,adLockReadOnly

        FTotalPage =  CLng(FTotalCount\FPageSize)
		if ((FTotalCount\FPageSize)<>(FTotalCount/FPageSize)) then
			FTotalPage = FtotalPage + 1
		end if
		FResultCount = db3_rsget.RecordCount
        if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not db3_rsget.EOF  then
            do until db3_rsget.eof
				set FItemList(i) = new CDeliverySongjangChangeLogItem

                FItemList(i).Fodetailidx        = db3_rsget("odetailidx")
                FItemList(i).Forderserial       = db3_rsget("orderserial")
                FItemList(i).Fpsongjangno       = db3_rsget("psongjangno")
                FItemList(i).Fpsongjangdiv      = db3_rsget("psongjangdiv")
                FItemList(i).Fchgsongjangno     = db3_rsget("chgsongjangno")
                FItemList(i).Fchgsongjangdiv    = db3_rsget("chgsongjangdiv")
                FItemList(i).Fchguserid         = db3_rsget("chguserid")
                FItemList(i).Fupddt             = db3_rsget("upddt")

                FItemList(i).Forderserial       = db3_rsget("orderserial")

                FItemList(i).Fitemid            = db3_rsget("itemid")
                FItemList(i).Fitemoption        = db3_rsget("itemoption")
                FItemList(i).Fitemno            = db3_rsget("itemno")

                FItemList(i).Fmakerid           = db3_rsget("makerid")
                FItemList(i).Fitemname          = db3_rsget("itemname")
                FItemList(i).Fitemoptionname    = db3_rsget("itemoptionname")

                FItemList(i).Fsongjangno        = db3_rsget("songjangno")
                FItemList(i).Fsongjangdiv       = db3_rsget("songjangdiv")
                FItemList(i).Fbeasongdate       = db3_rsget("beasongdate")
                FItemList(i).Fdlvfinishdt       = db3_rsget("dlvfinishdt")
                FItemList(i).Fjungsanfixdate    = db3_rsget("jungsanfixdate")

                FItemList(i).Fdcancelyn     = db3_rsget("dcancelyn")
                FItemList(i).Fsitename      = db3_rsget("sitename")
                FItemList(i).Fcancelyn      = db3_rsget("cancelyn")
                FItemList(i).Fauthcode      = db3_rsget("authcode")
                FItemList(i).Freqzipaddr    = db3_rsget("reqzipaddr")
                FItemList(i).Fcomment       = db3_rsget("comment")

				db3_rsget.moveNext
				i=i+1
			loop
        end if
        db3_rsget.close()
    end function
    
    public function getSongjangChangeLogList()
        Dim sqlStr
        sqlStr = "exec [db_order].[dbo].[usp_Ten_Delivery_Trace_ChgOrderSongjang_GetLOG] '"&FRectOrderSerial&"'"
 
        dbget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly

        FTotalPage =  CLng(FTotalCount\FPageSize)
		if ((FTotalCount\FPageSize)<>(FTotalCount/FPageSize)) then
			FTotalPage = FtotalPage + 1
		end if
		FResultCount = rsget.RecordCount
        if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
            do until rsget.eof
				set FItemList(i) = new CDeliverySongjangChangeLogItem
                FItemList(i).Fsongjangchgidx    = rsget("songjangchgidx")
                FItemList(i).Fodetailidx        = rsget("odetailidx")
                FItemList(i).Forderserial       = rsget("orderserial")
                FItemList(i).Fpsongjangno       = rsget("psongjangno")
                FItemList(i).Fpsongjangdiv      = rsget("psongjangdiv")
                FItemList(i).Fchgsongjangno     = rsget("chgsongjangno")
                FItemList(i).Fchgsongjangdiv    = rsget("chgsongjangdiv")
                FItemList(i).Fchguserid         = rsget("chguserid")
                FItemList(i).Fregdt             = rsget("regdt")
                FItemList(i).FactionType        = rsget("actionType")
                FItemList(i).Fchgdlvfinishdt    = rsget("chgdlvfinishdt")
                FItemList(i).Fitemid            = rsget("itemid")
                FItemList(i).Fitemoption        = rsget("itemoption")
                FItemList(i).Fsongjangno        = rsget("songjangno")
                FItemList(i).Fsongjangdiv       = rsget("songjangdiv")
                FItemList(i).Fbeasongdate       = rsget("beasongdate")
                FItemList(i).Fdlvfinishdt       = rsget("dlvfinishdt")
                FItemList(i).Fjungsanfixdate    = rsget("jungsanfixdate")

				rsget.moveNext
				i=i+1
			loop
        end if
        rsget.close()

    end function

    public function getDeliveryTrackExceptFinBrandList()
        Dim sqlStr, arrVal, i
        
        sqlStr = "exec [db_order].[dbo].[usp_TEN_Delivery_Trace_GetExceptBrandCNT] "&CHKIIF(FRectSongjangDiv="","NULL",FRectSongjangDiv)&",'"&FRectMakerid&"'"
        
        dbget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly
        if NOT rsget.Eof then
            FTotalCount = rsget("cnt")
        end if
        rsget.close()

        sqlStr = "exec [db_order].[dbo].[usp_TEN_Delivery_Trace_GetExceptBrandLIST] "&FCurrPage&","&FPageSize&","&CHKIIF(FRectSongjangDiv="","NULL",FRectSongjangDiv)&",'"&FRectMakerid&"'"
'rw sqlStr ': exit function
        dbget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly

        FTotalPage =  CLng(FTotalCount\FPageSize)
		if ((FTotalCount\FPageSize)<>(FTotalCount/FPageSize)) then
			FTotalPage = FtotalPage + 1
		end if
		FResultCount = rsget.RecordCount
        if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
            do until rsget.eof
				set FItemList(i) = new CCSDeliveryTrackDetailItem

                FItemList(i).Fmakerid       = rsget("makerid")
                FItemList(i).Fsongjangdiv   = rsget("songjangdiv")
                FItemList(i).Fregdt         = rsget("regdate")
                FItemList(i).Freguserid     = rsget("reguserid")
                FItemList(i).Fdivname       = rsget("divname")

				rsget.moveNext
				i=i+1
			loop
        end if
        rsget.close()

    end function

    public function getDeliveryTrackMifinListAdm()
        Dim sqlStr, arrVal, i

        'if NOT (FRectSearchType="1" or FRectSearchType="2" or FRectSearchType="9") then
            sqlStr = "exec [db_dataSummary].[dbo].[usp_Ten_Delivery_Track_MiBeasong_CNT]  '"&FRectStartDate&"','"&FRectEndDate&"',"&CHKIIF(FRectSongjangDiv="","NULL",FRectSongjangDiv)&",'"&FRectMakerid&"','"&FRectOrderserial&"',"&FRectSearchType&""

            db3_dbget.CursorLocation = adUseClient
            db3_rsget.Open sqlStr,db3_dbget,adOpenForwardOnly,adLockReadOnly
            if NOT db3_rsget.Eof then
                FTotalCount = db3_rsget("cnt")
            end if
            db3_rsget.close()
        'end if

        sqlStr = "exec [db_dataSummary].[dbo].[usp_Ten_Delivery_Track_MiBeasong_LIST] "&FCurrPage&","&FPageSize&",'"&FRectStartDate&"','"&FRectEndDate&"',"&CHKIIF(FRectSongjangDiv="","NULL",FRectSongjangDiv)&",'"&FRectMakerid&"','"&FRectOrderserial&"',"&FRectSearchType&""

        'response.write sqlStr & "<Br>"
        db3_dbget.CursorLocation = adUseClient
        db3_rsget.Open sqlStr,db3_dbget,adOpenForwardOnly,adLockReadOnly

        'if (FRectSearchType="1" or FRectSearchType="2" or FRectSearchType="9") then
        '    FTotalCount=db3_rsget.RecordCount
        '    FTotalPage = 1
        '    FResultCount = FTotalCount
        'else

            FTotalPage =  CLng(FTotalCount\FPageSize)
            if ((FTotalCount\FPageSize)<>(FTotalCount/FPageSize)) then
                FTotalPage = FtotalPage + 1
            end if
            FResultCount = db3_rsget.RecordCount
            if FResultCount<1 then FResultCount=0
        'end if

		redim preserve FItemList(FResultCount)
		i=0
		if  not db3_rsget.EOF  then
            do until db3_rsget.eof

                set FItemList(i) = new CCSDeliveryTrackDetailItem
                
                FItemList(i).Forderserial   = db3_rsget("orderserial")
                FItemList(i).Fsongjangno    = db3_rsget("songjangno")
                FItemList(i).Fsongjangdiv   = db3_rsget("songjangdiv")
                FItemList(i).Fmakerid       = db3_rsget("makerid")

                
                FItemList(i).Fbeasongdate   = db3_rsget("beasongdate")

                ' FItemList(i).FdlvfinishDT   = db3_rsget("dlvfinishDT")
                ' FItemList(i).Fjungsanfixdate   = db3_rsget("jungsanfixdate")
                

                ''FItemList(i).Fdeparturedt   = db3_rsget("departuredt")
                
                'FItemList(i).Flastupdate    = db3_rsget("lastupdate")

                'FItemList(i).Fquelastupddt  = db3_rsget("quelastupddt")
                'FItemList(i).Fquelastupdno  = db3_rsget("quelastupdno")

                FItemList(i).Ftrdeparturedt = db3_rsget("trdeparturedt")
                FItemList(i).Ftrarrivedt    = db3_rsget("trarrivedt")
                FItemList(i).Ftrupddt       = db3_rsget("trupddt")
                
                FItemList(i).Fdivname       = db3_rsget("divname")
                FItemList(i).FsongjangTrackURL    = db2HTML(db3_rsget("findurl"))

                'FItemList(i).Fbuyname       = db2html(db3_rsget("buyname"))
                FItemList(i).Freqname       = db2html(db3_rsget("reqname"))
                FItemList(i).Freqzipaddr    = db2html(db3_rsget("reqzipaddr"))
                FItemList(i).Fsitename      = db2html(db3_rsget("sitename"))
                
				db3_rsget.moveNext
				i=i+1
			loop
        end if
        db3_rsget.close()

    end function

    public function getFakeSongjangErrDlvListAdm()
        Dim sqlStr, arrVal, i

        sqlStr = "exec [db_order].[dbo].[usp_Ten_Delivery_Trace_GetFakeList_DlvErr_LIST] '"&FRectStartDate&"','"&FRectEndDate&"'"

        dbget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly

        FTotalPage =  1
		
		FResultCount = rsget.RecordCount
        if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
            do until rsget.eof
                set FItemList(i) = new CCSDeliveryTrackDetailItem
                FItemList(i).Fodetailidx    = rsget("odetailidx")
                FItemList(i).Forderserial   = rsget("orderserial")
                FItemList(i).Fsongjangno    = rsget("songjangno")
                FItemList(i).Fsongjangdiv   = rsget("songjangdiv")
                FItemList(i).Fmakerid       = rsget("makerid")
                FItemList(i).Fbeasongdate   = rsget("beasongdate")
                ''FItemList(i).Fdeparturedt   = rsget("departuredt")
                ''FItemList(i).FdlvfinishDT   = rsget("dlvfinishDT")
                FItemList(i).Flastupdate    = rsget("lastupdate")

                FItemList(i).Fquelastupddt  = rsget("quelastupddt")
                FItemList(i).Fquelastupdno  = rsget("quelastupdno")

                FItemList(i).Ftrdeparturedt = rsget("trdeparturedt")
                FItemList(i).Ftrarrivedt    = rsget("trarrivedt")
                FItemList(i).Ftrupddt       = rsget("trupddt")
                
                FItemList(i).Fdivname       = rsget("divname")
                FItemList(i).FsongjangTrackURL    = db2HTML(rsget("findurl"))

                FItemList(i).Fbuyname       = db2html(rsget("buyname"))
                FItemList(i).Freqname       = db2html(rsget("reqname"))
                FItemList(i).Freqzipaddr    = db2html(rsget("reqzipaddr"))
                FItemList(i).Fsitename      = db2html(rsget("sitename"))
                FItemList(i).Fitemid        = rsget("itemid")
                FItemList(i).Fitemname      = db2html(rsget("itemname"))
                FItemList(i).Fitemoptionname = db2html(rsget("itemoptionname"))



				rsget.moveNext
				i=i+1
			loop
        end if
        rsget.close()

    end function

    '' fake 송장 조회 업체용.
    public function getFakeSongjangErrDlvListForBrand()
        Dim sqlStr, arrVal, i

        if FRectMakerid="" then Exit function

        sqlStr = "exec [db_order].[dbo].[usp_Ten_Delivery_Trace_BrandView_GetFakeList_CNT]  '"&FRectStartDate&"','"&FRectEndDate&"','"&FRectMakerid&"'"

        dbget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly
        if NOT rsget.Eof then
            FTotalCount = rsget("cnt")
        end if
        rsget.close()

        sqlStr = "exec [db_order].[dbo].[usp_Ten_Delivery_Trace_BrandView_GetFakeList] "&FCurrPage&","&FPageSize&",'"&FRectStartDate&"','"&FRectEndDate&"','"&FRectMakerid&"'"

        dbget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly

        FTotalPage =  CLng(FTotalCount\FPageSize)
		if ((FTotalCount\FPageSize)<>(FTotalCount/FPageSize)) then
			FTotalPage = FtotalPage + 1
		end if
		FResultCount = rsget.RecordCount
        if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
            do until rsget.eof
                set FItemList(i) = new CCSDeliveryTrackDetailItem
                FItemList(i).Fodetailidx    = rsget("odetailidx")
                FItemList(i).Forderserial   = rsget("orderserial")
                FItemList(i).Fsongjangno    = rsget("songjangno")
                FItemList(i).Fsongjangdiv   = rsget("songjangdiv")
                FItemList(i).Fmakerid       = rsget("makerid")
                FItemList(i).Fbeasongdate   = rsget("beasongdate")
                ''FItemList(i).Fdeparturedt   = rsget("departuredt")
                ''FItemList(i).FdlvfinishDT   = rsget("dlvfinishDT")
                FItemList(i).Flastupdate    = rsget("lastupdate")

                FItemList(i).Fquelastupddt  = rsget("quelastupddt")
                FItemList(i).Fquelastupdno  = rsget("quelastupdno")

                FItemList(i).Ftrdeparturedt = rsget("trdeparturedt")
                FItemList(i).Ftrarrivedt    = rsget("trarrivedt")
                FItemList(i).Ftrupddt       = rsget("trupddt")
                
                FItemList(i).Fdivname       = rsget("divname")
                FItemList(i).FsongjangTrackURL    = db2HTML(rsget("findurl"))

                FItemList(i).Fbuyname       = db2html(rsget("buyname"))
                FItemList(i).Freqname       = db2html(rsget("reqname"))
                FItemList(i).Freqzipaddr    = db2html(rsget("reqzipaddr"))
                FItemList(i).Fitemid        = rsget("itemid")
                FItemList(i).Fitemname      = db2html(rsget("itemname"))
                FItemList(i).Fitemoptionname = db2html(rsget("itemoptionname"))

				rsget.moveNext
				i=i+1
			loop
        end if
        rsget.close()
    end function

    public function getFakeSongjangGrpBrandListAdm()
        Dim sqlStr, arrVal, i
        
        sqlStr = "exec [db_order].[dbo].[usp_Ten_Delivery_Trace_GetFakeListGrpByBrand_CNT]  '"&FRectStartDate&"','"&FRectEndDate&"',"&CHKIIF(FRectSongjangDiv="","NULL",FRectSongjangDiv)&",'"&FRectMakerid&"',"&FRectEtcdivinc&","&FRectByList&""

        dbget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly
        if NOT rsget.Eof then
            FTotalCount = rsget("cnt")
            if (FRectMakerid="") and (FRectByList="0") then
                FSumdelayTTL        = rsget("delayTTL")
                FSumMibeaTTL        = rsget("mibeaTTL")
                FSummijiphaTTL      = rsget("mijiphaTTL")
                FSumjiphaNoMoveTTL  = rsget("jiphaNoMoveTTL")

                FSumdelayTTLOrderGrp   = rsget("delayTTLOrderGrp")
                FSummibeaTTLOrderGrp   = rsget("mibeaTTLOrderGrp")
                FSummijiphaTTLOrderGrp = rsget("mijiphaTTLOrderGrp")
                FSumjiphaNoMoveTTLOrderGrp = rsget("jiphaNoMoveTTLOrderGrp")
            end if
        end if
        rsget.close()

        sqlStr = "exec [db_order].[dbo].[usp_Ten_Delivery_Trace_GetFakeListGrpByBrand_LIST] "&FCurrPage&","&FPageSize&",'"&FRectStartDate&"','"&FRectEndDate&"',"&CHKIIF(FRectSongjangDiv="","NULL",FRectSongjangDiv)&",'"&FRectMakerid&"',"&FRectEtcdivinc&","&FRectByList&""

        dbget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly

        FTotalPage =  CLng(FTotalCount\FPageSize)
		if ((FTotalCount\FPageSize)<>(FTotalCount/FPageSize)) then
			FTotalPage = FtotalPage + 1
		end if
		FResultCount = rsget.RecordCount
        if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
            do until rsget.eof
                if (FRectMakerid<>"") then      ' or (FRectByList<>"0")     프로시저에서는 뺀듯?
                    set FItemList(i) = new CCSDeliveryTrackDetailItem
                    if FRectMakerid<>"" and not(isnull(FRectMakerid)) then
                    FItemList(i).Fodetailidx    = rsget("odetailidx")
                    end if
                    FItemList(i).Forderserial   = rsget("orderserial")
                    FItemList(i).Fsongjangno    = rsget("songjangno")
                    FItemList(i).Fsongjangdiv   = rsget("songjangdiv")
                    FItemList(i).Fmakerid       = rsget("makerid")

                    
                    FItemList(i).Fbeasongdate   = rsget("beasongdate")

                    ''FItemList(i).Fdeparturedt   = rsget("departuredt")
                    ''FItemList(i).FdlvfinishDT   = rsget("dlvfinishDT")
                    FItemList(i).Flastupdate    = rsget("lastupdate")

                    FItemList(i).Fquelastupddt  = rsget("quelastupddt")
                    FItemList(i).Fquelastupdno  = rsget("quelastupdno")

                    FItemList(i).Ftrdeparturedt = rsget("trdeparturedt")
                    FItemList(i).Ftrarrivedt    = rsget("trarrivedt")
                    FItemList(i).Ftrupddt       = rsget("trupddt")
                    
                    FItemList(i).Fdivname       = rsget("divname")
                    FItemList(i).FsongjangTrackURL    = db2HTML(rsget("findurl"))

                    FItemList(i).Fbuyname       = db2html(rsget("buyname"))
                    FItemList(i).Freqname       = db2html(rsget("reqname"))
                    FItemList(i).Freqzipaddr    = db2html(rsget("reqzipaddr"))
                    FItemList(i).Fsitename      = db2html(rsget("sitename"))
                    FItemList(i).Fitemid        = rsget("itemid")
                    FItemList(i).Fitemname      = db2html(rsget("itemname"))
                    FItemList(i).Fitemoptionname = db2html(rsget("itemoptionname"))

                else
                    set FItemList(i) = new CCSDeliveryTrackBrandSumItem

                    FItemList(i).Fmakerid       = rsget("makerid")
                    FItemList(i).FTTLDLVFIN     = rsget("TTLDLVFIN")
                    FItemList(i).FdelayTTL      = rsget("delayTTL")
                    FItemList(i).FmibeaTTL      = rsget("mibeaTTL")
                    FItemList(i).FmijiphaTTL    = rsget("mijiphaTTL")
                    FItemList(i).FjiphaNoMoveTTL   = rsget("jiphaNoMoveTTL")

                    FItemList(i).FdelayTTLOrderGrp      = rsget("delayTTLOrderGrp")
                    FItemList(i).FmibeaTTLOrderGrp      = rsget("mibeaTTLOrderGrp")
                    FItemList(i).FmijiphaTTLOrderGrp    = rsget("mijiphaTTLOrderGrp")
                    FItemList(i).FjiphaNoMoveTTLOrderGrp   = rsget("jiphaNoMoveTTLOrderGrp")
                end if

				rsget.moveNext
				i=i+1
			loop
        end if
        rsget.close()

    end function

    public function getDeliveryStatusBrandListAdm()
        Dim sqlStr, arrVal, i
        
        sqlStr = "exec [db_dataSummary].[dbo].[usp_Ten_DeliveryListByBrand_CNT]  '"&FRectStartDate&"','"&FRectEndDate&"','"&FRectMakerid&"',"&CHKIIF(FRectSongjangDiv="","NULL",FRectSongjangDiv)&""

        db3_dbget.CursorLocation = adUseClient
        db3_rsget.Open sqlStr,db3_dbget,adOpenForwardOnly,adLockReadOnly
        if NOT db3_rsget.Eof then
            FTotalCount = db3_rsget("cnt")
            
        end if
        db3_rsget.close()

        sqlStr = "exec [db_dataSummary].[dbo].[usp_Ten_DeliveryListByBrand_LIST] "&FCurrPage&","&FPageSize&",'"&FRectStartDate&"','"&FRectEndDate&"','"&FRectMakerid&"',"&CHKIIF(FRectSongjangDiv="","NULL",FRectSongjangDiv)&""

        db3_dbget.CursorLocation = adUseClient
        db3_rsget.Open sqlStr,db3_dbget,adOpenForwardOnly,adLockReadOnly

        FTotalPage =  CLng(FTotalCount\FPageSize)
		if ((FTotalCount\FPageSize)<>(FTotalCount/FPageSize)) then
			FTotalPage = FtotalPage + 1
		end if
		FResultCount = db3_rsget.RecordCount
        if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not db3_rsget.EOF  then
            do until db3_rsget.eof
                set FItemList(i) = new CCSDeliveryTrackDetailItem
                FItemList(i).Fodetailidx    = db3_rsget("idx")
                FItemList(i).Forderserial   = db3_rsget("orderserial")
                FItemList(i).Fsongjangno    = db3_rsget("songjangno")
                FItemList(i).Fsongjangdiv   = db3_rsget("songjangdiv")
                FItemList(i).Fmakerid       = db3_rsget("makerid")
                FItemList(i).Fbeasongdate   = db3_rsget("beasongdate")
                FItemList(i).FdlvfinishDT   = db3_rsget("dlvfinishDT")
                FItemList(i).Fjungsanfixdate   = db3_rsget("jungsanfixdate")
                
                FItemList(i).Ftrdeparturedt = db3_rsget("trdeparturedt")
                FItemList(i).Ftrarrivedt    = db3_rsget("trarrivedt")
                FItemList(i).Ftrupddt       = db3_rsget("trupddt")
                
                FItemList(i).Fdivname       = db3_rsget("divname")
                FItemList(i).FsongjangTrackURL    = db2HTML(db3_rsget("findurl"))

                FItemList(i).Fbuyname       = db2html(db3_rsget("buyname"))
                FItemList(i).Freqname       = db2html(db3_rsget("reqname"))
                FItemList(i).Freqzipaddr    = db2html(db3_rsget("reqzipaddr"))
                FItemList(i).Fsitename      = db2html(db3_rsget("sitename"))
                FItemList(i).Fitemid        = db3_rsget("itemid")
                FItemList(i).Fitemname      = db2html(db3_rsget("itemname"))
                FItemList(i).Fitemoptionname = db2html(db3_rsget("itemoptionname"))

                FItemList(i).Fipkumdate     = db3_rsget("ipkumdate")
                FItemList(i).Fupcheconfirmdate   = db3_rsget("upcheconfirmdate")
                FItemList(i).Fipkumdiv      = db3_rsget("ipkumdiv")
                FItemList(i).Fjumundiv      = db3_rsget("jumundiv")
                FItemList(i).Fcurrstate     = db3_rsget("currstate")

				db3_rsget.moveNext
				i=i+1
			loop
        end if
        db3_rsget.close()

    end function

    public function getDeliveryTrackOrderInfo()
        Dim sqlStr, arrVal, i
        sqlStr = "exec [db_order].[dbo].[usp_Ten_Delivery_Trace_GetOrderInfo] '"&FRectOrderserial&"','"&FRectMakerid&"'"

        dbget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly
        if Not rsget.Eof then
            getDeliveryTrackOrderInfo = rsget.getRows()
        end if
        rsget.close()
    end function
    
    public function getDeliveryTrackOneInfo()
        Dim sqlStr, arrVal, i
        sqlStr = "exec [db_order].[dbo].[usp_Ten_Delivery_Trace_GetState_BySongjangNo] '"&FRectSongjangNo&"'"
        dbget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly

        FResultCount = rsget.RecordCount
        redim preserve FItemList(FResultCount)
		i=0

        if  not rsget.EOF  then
            
            do until rsget.eof
				set FItemList(i) = new CCSDeliveryTrackDetailItem
                FItemList(i).Ftracetbltype  = rsget("tbltype")
                FItemList(i).Fsongjangno    = rsget("songjangno")
                FItemList(i).Fsongjangdiv   = rsget("songjangdiv")
                FItemList(i).Fregdt         = rsget("regdt")
                FItemList(i).Fdeparturedt   = rsget("departuredt")
                FItemList(i).FdlvfinishDT   = rsget("arrivedt")
                FItemList(i).Ftraceupddt    = rsget("upddt")
                FItemList(i).FtraceAcctCnt  = rsget("lastupdno")

                rsget.moveNext
				i=i+1
			loop
        end if
        rsget.close()
    end function

    public function getDeliveryTrackSummaryDetailRealTime()
        Dim sqlStr, arrVal, i

        sqlStr = "exec [db_dataSummary].[dbo].[usp_TEN_delveryTrack_SummaryDetail_GETCNT] '"&FRectStartDate&"','"&FRectEndDate&"',"&CHKIIF(FRectSongjangDiv="","NULL",FRectSongjangDiv)&",'"&FRectMakerid&"','"&FRectIsUpchebeasong&"',"&FRectMijipHaExists&","&FRectMiBeasongExists&","&FRectEtcdivinc&","&CHKIIF(FRectErrChkType="",-1,FRectErrChkType)&",'"&FRectOrderserial&"','"&FRectSongjangNo&"'"
        
        db3_dbget.CursorLocation = adUseClient
        db3_rsget.Open sqlStr,db3_dbget,adOpenForwardOnly,adLockReadOnly
        if NOT db3_rsget.Eof then
            FTotalCount = db3_rsget("cnt")
        end if
        db3_rsget.close()

        sqlStr = "exec [db_order].[dbo].[usp_Ten_Delivery_Trace_SummaryDetail_ERR_GETLIST] "&FCurrPage&","&FPageSize&",'"&FRectStartDate&"','"&FRectEndDate&"',"&CHKIIF(FRectSongjangDiv="","NULL",FRectSongjangDiv)&",'"&FRectMakerid&"','"&FRectIsUpchebeasong&"',"&FRectMijipHaExists&","&FRectMiBeasongExists&","&FRectEtcdivinc&","&CHKIIF(FRectErrChkType="",-1,FRectErrChkType)&",'"&FRectOrderserial&"','"&FRectSongjangNo&"'"

        dbget.CursorLocation = adUseClient
        rsget.Open sqlStr,dbget,adOpenForwardOnly,adLockReadOnly

        FTotalPage =  CLng(FTotalCount\FPageSize)
		if ((FTotalCount\FPageSize)<>(FTotalCount/FPageSize)) then
			FTotalPage = FtotalPage + 1
		end if
		FResultCount = rsget.RecordCount
        if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not rsget.EOF  then
            do until rsget.eof
				set FItemList(i) = new CCSDeliveryTrackDetailItem

				FItemList(i).Forderserial   = rsget("orderserial")
                FItemList(i).Fsongjangno    = rsget("songjangno")
                FItemList(i).Fsongjangdiv   = rsget("songjangdiv")
                FItemList(i).Fmakerid       = rsget("makerid")
                FItemList(i).Fbeasongdate   = rsget("beasongdate")
                FItemList(i).Fdeparturedt   = rsget("departuredt")
                FItemList(i).FdlvfinishDT   = rsget("dlvfinishDT")
               '' FItemList(i).Ftraceupddt    = rsget("traceupddt")
                FItemList(i).Flastupdate    = rsget("lastupdate")
                FItemList(i).Fisupchebeasong= rsget("isupchebeasong")
                FItemList(i).Fdivname       = rsget("divname")
                FItemList(i).FsongjangTrackURL  = db2HTML(rsget("findurl"))

                FItemList(i).Fchktype       = rsget("chktype")

                FItemList(i).Ftrarrivedt   = rsget("arrivedt")
                FItemList(i).Ftraceupddt   = rsget("trupddt")

                FItemList(i).FcsCNT   = rsget("cs600CNT")
                FItemList(i).FcsFinCNT   = rsget("cs600FINCNT")

				rsget.moveNext
				i=i+1
			loop
        end if
        rsget.close()
    end function

    public function getDeliveryTrackSummaryDetail()
        Dim sqlStr, arrVal, i

        sqlStr = "exec [db_dataSummary].[dbo].[usp_TEN_delveryTrack_SummaryDetail_GETCNT] '"&FRectStartDate&"','"&FRectEndDate&"',"&CHKIIF(FRectSongjangDiv="","NULL",FRectSongjangDiv)&",'"&FRectMakerid&"','"&FRectIsUpchebeasong&"',"&FRectMijipHaExists&","&FRectMiBeasongExists&","&FRectEtcdivinc&","&CHKIIF(FRectErrChkType="",-1,FRectErrChkType)&",'"&FRectOrderserial&"','"&FRectSongjangNo&"'"
        
        db3_dbget.CursorLocation = adUseClient
        db3_rsget.Open sqlStr,db3_dbget,adOpenForwardOnly,adLockReadOnly
        if NOT db3_rsget.Eof then
            FTotalCount = db3_rsget("cnt")
        end if
        db3_rsget.close()

        sqlStr = "exec [db_dataSummary].[dbo].[usp_TEN_delveryTrack_SummaryDetail_GETLIST] "&FCurrPage&","&FPageSize&",'"&FRectStartDate&"','"&FRectEndDate&"',"&CHKIIF(FRectSongjangDiv="","NULL",FRectSongjangDiv)&",'"&FRectMakerid&"','"&FRectIsUpchebeasong&"',"&FRectMijipHaExists&","&FRectMiBeasongExists&","&FRectEtcdivinc&","&CHKIIF(FRectErrChkType="",-1,FRectErrChkType)&",'"&FRectOrderserial&"','"&FRectSongjangNo&"'"

        db3_dbget.CursorLocation = adUseClient
        db3_rsget.Open sqlStr,db3_dbget,adOpenForwardOnly,adLockReadOnly

        FTotalPage =  CLng(FTotalCount\FPageSize)
		if ((FTotalCount\FPageSize)<>(FTotalCount/FPageSize)) then
			FTotalPage = FtotalPage + 1
		end if
		FResultCount = db3_rsget.RecordCount
        if FResultCount<1 then FResultCount=0

		redim preserve FItemList(FResultCount)
		i=0
		if  not db3_rsget.EOF  then
            do until db3_rsget.eof
				set FItemList(i) = new CCSDeliveryTrackDetailItem

				FItemList(i).Forderserial   = db3_rsget("orderserial")
                FItemList(i).Fsongjangno    = db3_rsget("songjangno")
                FItemList(i).Fsongjangdiv   = db3_rsget("songjangdiv")
                FItemList(i).Fmakerid       = db3_rsget("makerid")
                FItemList(i).Fbeasongdate   = db3_rsget("beasongdate")
                FItemList(i).Fdeparturedt   = db3_rsget("departuredt")
                FItemList(i).FdlvfinishDT   = db3_rsget("dlvfinishDT")
                FItemList(i).Ftraceupddt    = db3_rsget("traceupddt")
                FItemList(i).Flastupdate    = db3_rsget("lastupdate")
                FItemList(i).Fisupchebeasong= db3_rsget("isupchebeasong")
                FItemList(i).Fdivname       = db3_rsget("divname")
                FItemList(i).FsongjangTrackURL  = db2HTML(db3_rsget("findurl"))

                FItemList(i).Fchktype       = db3_rsget("chktype")

				db3_rsget.moveNext
				i=i+1
			loop
        end if
        db3_rsget.close()
    end function
 
    public function getDeliveryTrackSummary()
        Dim sqlStr, arrVal, i
        '' @stdt, @eddt, @showbeasongdate, @showsongjangdiv, @showbrand, @songjangdiv, @makerid, @isupchebeasong, @ExistsMiJipha, @ExistsMiBeasong
        
        sqlStr = "exec [db_dataSummary].[dbo].[usp_TEN_delveryTrack_Summary_GETLIST] '"&FRectStartDate&"','"&FRectEndDate&"',"&FRectGrpBeasongdate&","&FRectGrpSongjangDiv&","&FRectGrpBrand&","&CHKIIF(FRectSongjangDiv="","NULL",FRectSongjangDiv)&",'"&FRectMakerid&"','"&FRectIsUpchebeasong&"',"&FRectMijipHaExists&","&FRectMiBeasongExists&","&FRectEtcdivinc&","&CHKIIF(FRectErrChkType="",-1,FRectErrChkType)
''rw sqlStr ': exit function

        db3_dbget.CursorLocation = adUseClient
        db3_rsget.Open sqlStr,db3_dbget,adOpenForwardOnly,adLockReadOnly

        FResultCount = db3_rsget.RecordCount
        FTotalCount = FResultCount

        redim preserve FItemList(FResultCount)
		i=0

        If not db3_rsget.EOF Then
            do until db3_rsget.eof
				set FItemList(i) = new CDeliveryTrackSmrItem
				FItemList(i).Fbeasongdate	    = db3_rsget("beasongdate")
				FItemList(i).Fsongjangdiv	    = db3_rsget("songjangdiv")
				FItemList(i).Fmakerid		    = db3_rsget("makerid")
                FItemList(i).Fisupchebeasong    = db3_rsget("isupchebeasong")
                FItemList(i).Fttlchulgono       = db3_rsget("ttlchulgono")
                FItemList(i).FjiphafinCNT       = db3_rsget("jiphafinCNT")
                FItemList(i).FdlvfinCNT         = db3_rsget("dlvfinCNT")
                FItemList(i).FDminusCnt         = db3_rsget("DminusCnt")
                FItemList(i).FDplus0Cnt         = db3_rsget("Dplus0Cnt")
                FItemList(i).FDplus1Cnt         = db3_rsget("Dplus1Cnt")
                FItemList(i).FDplus2Cnt         = db3_rsget("Dplus2Cnt")
                FItemList(i).FDplus3UpCnt       = db3_rsget("Dplus3UpCnt")
                FItemList(i).FMijiphaCnt        = db3_rsget("MijiphaCnt")
                FItemList(i).FMidlvfinCnt       = db3_rsget("MidlvfinCnt")
                FItemList(i).FsongjangDivName   = db3_rsget("divname")
                'FItemList(i).FsongjangTrackURL  = db3_rsget("songjangTrackURL")
                FItemList(i).FErrChkCnt         = db3_rsget("errchkcnt")

				db3_rsget.moveNext
				i=i+1
			loop
        End If
        db3_rsget.Close


    end function

    Private Sub Class_Initialize()
        FCurrPage       = 1
        FPageSize       = 10
        FResultCount    = 0
        FScrollCount    = 10
        FTotalCount     = 0
    End Sub

    Private Sub Class_Terminate()
		'//
    End Sub

    public Function HasPreScroll()
        HasPreScroll = StartScrollPage > 1
    end Function

    public Function HasNextScroll()
        HasNextScroll = FTotalPage > StartScrollPage + FScrollCount -1
    end Function

    public Function StartScrollPage()
        StartScrollPage = ((FCurrpage-1)\FScrollCount)*FScrollCount +1
    end Function
End Class



%>
