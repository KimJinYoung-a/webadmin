<%
'' include virtual="/admin/etc/incOutMallCommonFunction.asp  이쪽으로 합침

function TenDlvCode2InterParkDlvCode(itenCode)
    select Case itenCode
        CASE "1" : TenDlvCode2InterParkDlvCode = "169178"     ''한진
        CASE "2" : TenDlvCode2InterParkDlvCode = "169198"     ''현대
        CASE "3" : TenDlvCode2InterParkDlvCode = "169177"     ''대한통운
        CASE "4" : TenDlvCode2InterParkDlvCode = "169168"     ''CJ GLS
        CASE "5" : TenDlvCode2InterParkDlvCode = "169211"     ''이클라인
        CASE "6" : TenDlvCode2InterParkDlvCode = "169181"     ''삼성 HTH
        CASE "7" : TenDlvCode2InterParkDlvCode = "231145"     ''동부(구훼미리)
        CASE "8" : TenDlvCode2InterParkDlvCode = "169199"     ''우체국택배
        CASE "9" : TenDlvCode2InterParkDlvCode = "169187"     ''KGB택배
        CASE "10" : TenDlvCode2InterParkDlvCode = "169194"     ''아주택배 / 로엑스(구 아주)
        CASE "11" : TenDlvCode2InterParkDlvCode = ""     ''오렌지택배
        CASE "12" : TenDlvCode2InterParkDlvCode = ""     ''한국택배 / 뉴한국택배물류?
        CASE "13" : TenDlvCode2InterParkDlvCode = "169200"     ''옐로우캡
        CASE "14" : TenDlvCode2InterParkDlvCode = ""     ''나이스택배
        CASE "15" : TenDlvCode2InterParkDlvCode = ""     ''중앙택배
        CASE "16" : TenDlvCode2InterParkDlvCode = ""     ''주코택배
        CASE "17" : TenDlvCode2InterParkDlvCode = ""     ''트라넷택배
        CASE "18" : TenDlvCode2InterParkDlvCode = "169182"     ''로젠택배
        CASE "19" : TenDlvCode2InterParkDlvCode = ""     ''KGB특급택배
        CASE "20" : TenDlvCode2InterParkDlvCode = ""     ''KT로지스
        CASE "21" : TenDlvCode2InterParkDlvCode = "303978"     ''경동택배
        CASE "22" : TenDlvCode2InterParkDlvCode = "169526"     ''고려택배
        CASE "23" : TenDlvCode2InterParkDlvCode = "236288"     ''쎄덱스택배 신세계
        CASE "24" : TenDlvCode2InterParkDlvCode = "231491"     ''사가와익스프레스
        CASE "25" : TenDlvCode2InterParkDlvCode = "229381"     ''하나로택배
        CASE "26" : TenDlvCode2InterParkDlvCode = "263792"     ''일양택배
        CASE "27" : TenDlvCode2InterParkDlvCode = "169194"     ''LOEX택배
        CASE "28" : TenDlvCode2InterParkDlvCode = "231145"     ''동부익스프레스
        CASE "29" : TenDlvCode2InterParkDlvCode = "231194"     ''건영택배
        CASE "30" : TenDlvCode2InterParkDlvCode = "266237"     ''이노지스
        CASE "31" : TenDlvCode2InterParkDlvCode = "230175"     ''천일택배
        CASE "33" : TenDlvCode2InterParkDlvCode = "250701"     ''호남택배
        CASE "34" : TenDlvCode2InterParkDlvCode = "258064"     ''대신화물택배
        CASE "35" : TenDlvCode2InterParkDlvCode = "169172"     ''CVSnet택배
        CASE "98" : TenDlvCode2InterParkDlvCode = "169316"     ''퀵서비스->직배송
        CASE "99" : TenDlvCode2InterParkDlvCode = "169167"     ''기타
        CASE  Else
            TenDlvCode2InterParkDlvCode = ""      ''기타발송(169167)
    end Select
end function

function TenDlvCode2LotteDlvCode(itenCode)
    ''if IsNULL(itenCode) then Exit function
    if IsNULL(itenCode) then itenCode="99"
    
    itenCode = TRIM(CStr(itenCode))
    select Case itenCode
        CASE "1" : TenDlvCode2LotteDlvCode = "27"     ''한진
        CASE "2" : TenDlvCode2LotteDlvCode = "1"     ''현대v
        CASE "3" : TenDlvCode2LotteDlvCode = "5"     ''대한통운
        CASE "4" : TenDlvCode2LotteDlvCode = "31"     ''CJ GLS
        CASE "5" : TenDlvCode2LotteDlvCode = "23"     ''이클라인
        CASE "6" : TenDlvCode2LotteDlvCode = "32"     ''삼성 HTH
        CASE "7" : TenDlvCode2LotteDlvCode = "56"     ''동부(구훼미리) ''확
        CASE "8" : TenDlvCode2LotteDlvCode = "9339"     ''우체국택배
        CASE "9" : TenDlvCode2LotteDlvCode = "39"     ''KGB택배
        CASE "10" : TenDlvCode2LotteDlvCode = "34"     ''아주택배 / 로엑스(구 아주)
        CASE "11" : TenDlvCode2LotteDlvCode = ""     ''오렌지택배
        CASE "12" : TenDlvCode2LotteDlvCode = "29"     ''한국택배 / 한국특송
        CASE "13" : TenDlvCode2LotteDlvCode = "37"     ''옐로우캡
        CASE "14" : TenDlvCode2LotteDlvCode = ""     ''나이스택배
        CASE "15" : TenDlvCode2LotteDlvCode = "43"     ''중앙택배
        CASE "16" : TenDlvCode2LotteDlvCode = ""     ''주코택배
        CASE "17" : TenDlvCode2LotteDlvCode = "36"     ''트라넷택배
        CASE "18" : TenDlvCode2LotteDlvCode = "41"     ''로젠택배
        CASE "19" : TenDlvCode2LotteDlvCode = "44"     ''KGB특급택배
        CASE "20" : TenDlvCode2LotteDlvCode = "30"     ''KT로지스
        CASE "21" : TenDlvCode2LotteDlvCode = "52"     ''경동택배
        CASE "22" : TenDlvCode2LotteDlvCode = ""     ''고려택배
        CASE "23" : TenDlvCode2LotteDlvCode = "42"     ''쎄덱스택배 신세계
        CASE "24" : TenDlvCode2LotteDlvCode = "51"     ''사가와익스프레스
        CASE "25" : TenDlvCode2LotteDlvCode = "3"     ''하나로택배v
        CASE "26" : TenDlvCode2LotteDlvCode = "47"     ''일양택배
        CASE "27" : TenDlvCode2LotteDlvCode = ""     ''LOEX택배
        CASE "28" : TenDlvCode2LotteDlvCode = "35"     ''동부익스프레스
        CASE "29" : TenDlvCode2LotteDlvCode = "45"     ''건영택배
        CASE "30" : TenDlvCode2LotteDlvCode = "57"     ''이노지스
        CASE "31" : TenDlvCode2LotteDlvCode = "33"     ''천일택배
        CASE "33" : TenDlvCode2LotteDlvCode = "99"     ''호남택배
        CASE "34" : TenDlvCode2LotteDlvCode = "46"     ''대신화물택배
        CASE "35" : TenDlvCode2LotteDlvCode = "99"     ''CVSnet택배
        CASE "98" : TenDlvCode2LotteDlvCode = "99"     ''퀵서비스
        CASE "99" : TenDlvCode2LotteDlvCode = "99"     ''업체직송
        CASE  Else
            TenDlvCode2LotteDlvCode = "99"
    end Select
end function


'''롯데iMall 송장변환
function TenDlvCode2LotteiMallDlvCode(itenCode)
    if IsNULL(itenCode) then Exit function
    itenCode = TRIM(CStr(itenCode))
''41	이젠택배
''99	기타

    select Case itenCode
        CASE "1" : TenDlvCode2LotteiMallDlvCode = "15"     ''한진
        CASE "2" : TenDlvCode2LotteiMallDlvCode = "11"     ''현대v
        CASE "3" : TenDlvCode2LotteiMallDlvCode = "12"     ''대한통운
        CASE "4" : TenDlvCode2LotteiMallDlvCode = "16"     ''CJ GLS
        CASE "5" : TenDlvCode2LotteiMallDlvCode = ""     ''이클라인
        CASE "6" : TenDlvCode2LotteiMallDlvCode = "22"     ''삼성 HTH
        CASE "7" : TenDlvCode2LotteiMallDlvCode = "26"     ''동부(구훼미리) ''확
        CASE "8" : TenDlvCode2LotteiMallDlvCode = "31"     ''우체국택배
        CASE "9" : TenDlvCode2LotteiMallDlvCode = "40"     ''KGB택배
        CASE "10" : TenDlvCode2LotteiMallDlvCode = "34"     ''아주택배 / 로엑스(구 아주)
        CASE "11" : TenDlvCode2LotteiMallDlvCode = ""     ''오렌지택배
        CASE "12" : TenDlvCode2LotteiMallDlvCode = "37"     ''한국택배 / 한국특송
        CASE "13" : TenDlvCode2LotteiMallDlvCode = "32"     ''옐로우캡
        CASE "14" : TenDlvCode2LotteiMallDlvCode = ""     ''나이스택배
        CASE "15" : TenDlvCode2LotteiMallDlvCode = ""     ''중앙택배
        CASE "16" : TenDlvCode2LotteiMallDlvCode = ""     ''주코택배
        CASE "17" : TenDlvCode2LotteiMallDlvCode = "36"     ''트라넷택배
        CASE "18" : TenDlvCode2LotteiMallDlvCode = "24"     ''로젠택배
        CASE "19" : TenDlvCode2LotteiMallDlvCode = "40"     ''KGB특급택배
        CASE "20" : TenDlvCode2LotteiMallDlvCode = ""     ''KT로지스
        CASE "21" : TenDlvCode2LotteiMallDlvCode = "49"     ''경동택배
        CASE "22" : TenDlvCode2LotteiMallDlvCode = ""     ''고려택배
        CASE "23" : TenDlvCode2LotteiMallDlvCode = "47"     ''쎄덱스택배 신세계
        CASE "24" : TenDlvCode2LotteiMallDlvCode = "43"     ''사가와익스프레스
        CASE "25" : TenDlvCode2LotteiMallDlvCode = "46"     ''하나로택배v
        CASE "26" : TenDlvCode2LotteiMallDlvCode = "18"     ''일양택배
        CASE "27" : TenDlvCode2LotteiMallDlvCode = "48"     ''LOEX택배
        CASE "28" : TenDlvCode2LotteiMallDlvCode = "26"     ''동부익스프레스
        CASE "29" : TenDlvCode2LotteiMallDlvCode = "99"     ''건영택배
        CASE "30" : TenDlvCode2LotteiMallDlvCode = "23"     ''이노지스
        CASE "31" : TenDlvCode2LotteiMallDlvCode = "17"     ''천일택배
        CASE "33" : TenDlvCode2LotteiMallDlvCode = ""     ''호남택배
        CASE "34" : TenDlvCode2LotteiMallDlvCode = "38"     ''대신화물택배
        CASE "35" : TenDlvCode2LotteiMallDlvCode = "99"     ''CVSnet택배
        CASE "98" : TenDlvCode2LotteiMallDlvCode = "99"     ''퀵서비스
        CASE "99" : TenDlvCode2LotteiMallDlvCode = "99"     ''업체직송
        CASE  Else
            TenDlvCode2LotteiMallDlvCode = "99"
    end Select
end function

function LotteiMallDlvCode2Name(iltDlvCode)
    LotteiMallDlvCode2Name = "기타"
    if IsNULL(iltDlvCode) then Exit function
    iltDlvCode = TRIM(CStr(iltDlvCode))
    
    select Case iltDlvCode
        CASE "11" : LotteiMallDlvCode2Name="현대택배"
        CASE "12" : LotteiMallDlvCode2Name="씨제이대한통운"
        CASE "15" : LotteiMallDlvCode2Name="한진택배"
        CASE "16" : LotteiMallDlvCode2Name="CJGLS"
        CASE "17" : LotteiMallDlvCode2Name="천일택배"
        CASE "18" : LotteiMallDlvCode2Name="일양택배"
        CASE "19" : LotteiMallDlvCode2Name="기타택배"
        CASE "22" : LotteiMallDlvCode2Name="HTH택배"
        CASE "24" : LotteiMallDlvCode2Name="로젠택배"
        CASE "26" : LotteiMallDlvCode2Name="동부익스프레스"
        CASE "31" : LotteiMallDlvCode2Name="우체국택배"
        CASE "32" : LotteiMallDlvCode2Name="옐로우캡"
        CASE "34" : LotteiMallDlvCode2Name="아주택배"
        CASE "36" : LotteiMallDlvCode2Name="트라넷"
        CASE "37" : LotteiMallDlvCode2Name="한국택배"
        CASE "38" : LotteiMallDlvCode2Name="대신택배"
        CASE "40" : LotteiMallDlvCode2Name="KGB택배"
        CASE "41" : LotteiMallDlvCode2Name="이젠택배"
        CASE "43" : LotteiMallDlvCode2Name="사가와익스프레스"
        CASE "46" : LotteiMallDlvCode2Name="하나로택배"
        CASE "47" : LotteiMallDlvCode2Name="세덱스택배"
        CASE "48" : LotteiMallDlvCode2Name="로엑스택배"
        CASE "49" : LotteiMallDlvCode2Name="경동택배"
        CASE "99" : LotteiMallDlvCode2Name="기타"
        CASE  Else
            LotteiMallDlvCode2Name = "기타"
    end Select
end function

%>