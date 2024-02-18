<%

function getTopMenuJSon_TEST
    ''TODAY, LIVING , FASHION , LIFE STYLE , DIGITAL , BABY
    
    Dim URIFIX 
    URIFIX = "https://webadmin.10x10.co.kr"
    IF application("Svr_Info")="Dev" THEN
         URIFIX = "http://testwebadmin.10x10.co.kr"   
    end if
    
    dim objRst
    Set objRst = jsArray()

    SET objRst(Null) = jsObject()
    objRst(Null)("topid") ="dashboard"
    objRst(Null)("topname") = "게시판"
    objRst(Null)("topurl")  = ""
    objRst(Null)("topdefault")=1
	
	SET objRst(Null) = jsObject()
    objRst(Null)("topid") ="message"
    objRst(Null)("topname") = "알림"
    objRst(Null)("topurl")  = URIFIX & "/apps/academy/notice/pushList.asp"
    objRst(Null)("topdefault")=0
	
	SET objRst(Null) = jsObject()
    objRst(Null)("topid") ="itemmaster"
    objRst(Null)("topname") = "작품관리"
    objRst(Null)("topurl")  = URIFIX & "/apps/academy/itemmaster/artList.asp"
    objRst(Null)("topdefault")=0

	SET objRst(Null) = jsObject()
    objRst(Null)("topid") ="ordermaster"
    objRst(Null)("topname") = "주문관리"
    objRst(Null)("topurl")  = URIFIX & "/apps/academy/ordermaster/orderList.asp"
    objRst(Null)("topdefault")=0

	SET objRst(Null) = jsObject()
	objRst(Null)("topid") ="etc"
	objRst(Null)("topname") = "기타"
	objRst(Null)("topurl")  = URIFIX & "/apps/academy/etc/etcIndex.asp"
	objRst(Null)("topdefault")=0

    SET getTopMenuJSon_TEST = objRst
    
    Set objRst = Nothing
end Function

function getSubMenuJSon_TEST
    ''TODAY, LIVING , FASHION , LIFE STYLE , DIGITAL , BABY
    
    Dim URIFIX 
    URIFIX = "https://webadmin.10x10.co.kr"
    IF application("Svr_Info")="Dev" THEN
         URIFIX = "http://testwebadmin.10x10.co.kr"   
    end if
    
    dim objRst
    Set objRst = jsArray()

	SET objRst(Null) = jsObject()
	objRst(Null)("subid") ="cheer"
	objRst(Null)("suburl")  = URIFIX & "/apps/academy/etc/talk.asp"

	SET objRst(Null) = jsObject()
	objRst(Null)("subid") ="review"
	objRst(Null)("suburl")  = URIFIX & "/apps/academy/etc/reviewList.asp"

	SET objRst(Null) = jsObject()
	objRst(Null)("subid") ="qna"
	objRst(Null)("suburl")  = URIFIX & "/apps/academy/etc/QnaList.asp"

    SET getSubMenuJSon_TEST = objRst
    
    Set objRst = Nothing
end Function

function getFingerBoardMenuJSon
   
    Dim URIFIX 
    URIFIX = "https://webadmin.10x10.co.kr"
    IF application("Svr_Info")="Dev" THEN
         URIFIX = "http://testwebadmin.10x10.co.kr"   
    end if
    
    dim objRst
    Set objRst = jsArray()

	SET objRst(Null) = jsObject()
	objRst(Null)("subid") ="order"
	objRst(Null)("suburl")  = URIFIX & "/apps/academy/ordermaster/OrderList.asp?odiv=S"

	SET objRst(Null) = jsObject()
	objRst(Null)("subid") ="deliver"
	objRst(Null)("suburl")  = URIFIX & "/apps/academy/ordermaster/OrderList.asp?odiv=D"

	SET objRst(Null) = jsObject()
	objRst(Null)("subid") ="cs"
	objRst(Null)("suburl")  = URIFIX & "/apps/academy/ordermaster/OrderList.asp?odiv=C"

	SET objRst(Null) = jsObject()
	objRst(Null)("subid") ="qna"
	objRst(Null)("suburl")  = URIFIX & "/apps/academy/etc/QnaList.asp"

    SET getFingerBoardMenuJSon = objRst
    
    Set objRst = Nothing
end Function

function getTopMenuJSon_2015
    ''TODAY, LIVING , FASHION , LIFE STYLE , DIGITAL , BABY
    
    Dim URIFIX 
    URIFIX = "https://webadmin.10x10.co.kr"
    IF application("Svr_Info")="Dev" THEN
         URIFIX = "http://testwebadmin.10x10.co.kr"   
    end if
    
    dim spacingChar
    if InStr(Lcase(Request.ServerVariables("HTTP_USER_AGENT")),"ios")>1 then
        spacingChar = " "
    end if 
    dim objRst
    Set objRst = jsArray()

	SET objRst(Null) = jsObject()
    objRst(Null)("topid") ="dashboard"
    objRst(Null)("topname") = "알림판"
    objRst(Null)("topurl")  = ""
    objRst(Null)("topdefault") =1

	SET objRst(Null) = jsObject()
    objRst(Null)("topid") ="message"
    objRst(Null)("topname") = "메세지"
    objRst(Null)("topurl")  = URIFIX&"/apps/academy/notice/pushList.asp"
    objRst(Null)("topdefault") =0
	
	SET objRst(Null) = jsObject()
    objRst(Null)("topid") ="itemmaster"
    objRst(Null)("topname") = "작품관리"
    objRst(Null)("topurl")  = URIFIX&"/apps/academy/itemmaster/artList.asp"
    objRst(Null)("topdefault") =0
	
	SET objRst(Null) = jsObject()
	objRst(Null)("topid") ="profile"
	objRst(Null)("topname") = "설정"
	objRst(Null)("topurl")  = ""
	objRst(Null)("topdefault") =0
   
    SET getTopMenuJSon_2015 = objRst
    
    Set objRst = Nothing
end Function



function getRealTimeBestKeywords
    dim objRst
    DIM oPpkDoc, arrList, iRows
    Set objRst = jsArray()
    
    
	SET oPpkDoc = NEW SearchItemCls
		arrList = oPpkDoc.getPopularKeyWords()
	SET oPpkDoc = NOTHING 
	 
	IF isArray(arrList)  THEN	 
		if Ubound(arrList)>0 then
		    FOR iRows =0 To UBOUND(arrList)
		        Set objRst(Null) = jsObject()
    	        objRst(null)("keyword") = cStr(arrList(iRows))	
	        Next
	    END IF
    END IF            
						    
        
    SET getRealTimeBestKeywords = objRst
    
    Set objRst = Nothing
end function	

''순위변동 포함
function getRealTimeBestKeywords2
    dim objRst
    DIM oPpkDoc, arrList, iRows
    DIM arTg
    Set objRst = jsArray()
    
    
	SET oPpkDoc = NEW SearchItemCls
		Call oPpkDoc.getPopularKeyWords2(arrList, arTg)
	SET oPpkDoc = NOTHING 
	 
	IF isArray(arrList)  THEN	 
		if Ubound(arrList)>0 then
		    FOR iRows =0 To UBOUND(arrList)
		        Set objRst(Null) = jsObject()
    	        objRst(null)("keyword") = cStr(arrList(iRows))	
    	        if cStr(arTg(iRows))="" then
    	            objRst(null)("rank") = cStr("0")
    	        else
    	            objRst(null)("rank") = cStr(arTg(iRows))
    	        end if
	        Next
	    END IF
    END IF            
    SET getRealTimeBestKeywords2 = objRst
    Set objRst = Nothing
end function	

''검색어 자동완성
function getAutoCompleteKeywords(seed_str)
    dim objRst
    DIM oPpkDoc, ikwd_count, ikwd_list, icnv_str, iRows
    Set objRst = jsArray()
    
    
	SET oPpkDoc = NEW SearchItemCls
		Call oPpkDoc.getAutoCompleteKeywords(seed_str, ikwd_count, ikwd_list, icnv_str)
	SET oPpkDoc = NOTHING 
	 
	
    FOR iRows =0 To ikwd_count-1
        Set objRst(Null) = jsObject()
        objRst(null)("Word") = cStr(ikwd_list(iRows))	
        objRst(null)("Seed") = cStr(seed_str)	
        objRst(null)("Conv") = "" 'cStr(icnv_str)	search4에서 구조가 좀 바뀌었음  2015/03/09
    Next
						    
        
    SET getAutoCompleteKeywords = objRst
    
    Set objRst = Nothing
end function

''V3
function getLnbBgColorType(iuserid)
    getLnbBgColorType = 1  '' 1,2 
end function



'V3.1
function getLnbBottomJson()
    Dim URIFIX 
    URIFIX = "http://m.10x10.co.kr"
    IF application("Svr_Info")="Dev" THEN
         URIFIX = "http://testm.10x10.co.kr"   
    end if
    
	Dim IMGPREFIX : IMGPREFIX = "http://fiximage.10x10.co.kr/m/2016/common/"
    dim objRst : Set objRst = jsArray()
    Set objRst(Null) = jsObject()
    objRst(Null)("navname")= "PLAY"
    objRst(Null)("navicon")= IMGPREFIX&"/ico_play.png"
    objRst(Null)("navlink")= URIFIX&"/apps/academy/play/index.asp"
    
    Set objRst(Null) = jsObject()
    objRst(Null)("navname")= "WISH"
    objRst(Null)("navicon")= IMGPREFIX&"/ico_wish.png"
    objRst(Null)("navlink")= URIFIX&"/apps/academy/wish/index.asp"
    
    Set objRst(Null) = jsObject()
    objRst(Null)("navname")= "GIFT"
    objRst(Null)("navicon")= IMGPREFIX&"/ico_gift.png"
    objRst(Null)("navlink")= URIFIX&"/apps/academy/gift/gifttalk/index.asp"
    
    SET getLnbBottomJson = objRst
    
    SET objRst = Nothing
End Function
%>