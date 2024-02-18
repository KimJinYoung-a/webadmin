<%@  codepage="65001" language="VBScript" %>
<% option explicit %>
<%
session.codePage = 65001
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
Response.AddHeader "Content-Disposition","attachment;filename=이벤트_상품리스트_" & date & hour(now) & minute(now) & ".xls"
Response.ContentType = "application/vnd.ms-excel"
Response.CacheControl = "public"
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function_utf8.asp"-->
<!-- #include virtual="/lib/classes/event/eventManageCls_V5.asp"-->
<!-- #include virtual="/admin/eventmanage/common/event_function.asp"-->
<!-- #include virtual="/lib/classes/displaycate/displaycateCls.asp"-->
<!-- #include virtual="/admin/lib/incPageFunction.asp" -->
<%
Dim eCode
Dim cEvtItem,cEvtCont,cEGroup
Dim ekind,eman,escope,ename,esday,eeday,epday, elevel,estate,eregdate,estatedesc, ekinddesc
Dim arrGroup,arrGroup_mo
Dim iTotCnt, arrList,intLoop
Dim iPageSize, iCurrpage ,iDelCnt
Dim iStartPage, iEndPage, iTotalPage, ix,iPerCnt	
Dim iDispYCnt, iDispNCnt

Dim strG, strSort
Dim sDate,sSdate,sEdate, sEvt,strTxt, sCategory,sState,sKind,eisort
Dim strparm, Brand
dim makerid, itemname, itemid
dim blnOnlyMobile	,itemsort,eItemListType, blnWeb, blnMobile, blnApp
dim eChannel
dim arrList_mo, iTotCnt_mo, iTotalPage_mo
Dim iStartPage_mo, iEndPage_mo , ix_mo,iCurrpage_mo 
dim strG_mo, 	itemsort_mo
dim arrGroupMo, dispCate
  
strG  		= requestCheckvar(Request("selG"),10)
strSort  	= requestCheckvar(Request("selSort"),1)
	
eCode 		= requestCheckvar(request("eC"),10)
itemid      = requestCheckvar(request("itemid"),255)
itemname    = requestCheckvar(request("itemname"),64)
makerid     = requestCheckvar(request("makerid"),32)
itemsort  	= requestCheckvar(request("itemsort"),32)
eChannel    = requestCheckvar(request("eCh"),1)
dispCate = requestCheckvar(request("disp"),16)

if eChannel = "" then eChannel = "P"
	if itemid<>"" then
	dim iA ,arrTemp,arrItemid
	itemid = replace(itemid,",",chr(10))
	itemid = replace(itemid,chr(13),"")
	arrTemp = Split(itemid,chr(10))

	iA = 0
	do while iA <= ubound(arrTemp) 
		if trim(arrTemp(iA))<>"" then
			'상품코드 유효성 검사(2008.08.05;허진원)
			if Not(isNumeric(trim(arrTemp(iA)))) then
				Response.Write "<script language=javascript>alert('[" & arrTemp(iA) & "]은(는) 유효한 상품코드가 아닙니다.');history.back();</script>"
				dbget.close()	:	response.End
			else
				arrItemid = arrItemid & trim(arrTemp(iA)) & ","
			end if
		end if
		iA = iA + 1
	loop
	itemid = left(arrItemid,len(arrItemid)-1)
end if

	'파라미터값 받기 & 기본 변수 값 세팅
	iCurrpage = Request("iC")	'현재 페이지 번호
     
	IF iCurrpage = "" THEN
		iCurrpage = 1	
	END IF	  
	
	iCurrpage_mo = Request("iC_mo")	'현재 페이지 번호
     
	IF iCurrpage_mo = "" THEN
		iCurrpage_mo = 1	
	END IF	  
		
	iPageSize = 3000		'한 페이지의 보여지는 열의 수
	iPerCnt = 10		'보여지는 페이지 간격
	
	'## 검색 #############################			
	sDate = Request("selDate")  '기간 
	sSdate = Request("iSD")
	sEdate = Request("iED")	
	
	sEvt = Request("selEvt")  '이벤트 코드/명 검색
	strTxt = Request("sEtxt")
	
	sCategory	= Request("selC") '카테고리
	sState	 = Request("eventstate")'이벤트 상태	
	sKind = Request("eventkind")	'이벤트종류
 
	if blnOnlyMobile ="" then blnOnlyMobile = 0
	strparm = "selDate="&sDate&"&iSD="&sSdate&"&iED="&sEdate&"&selEvt="&sEvt&"&sEtxt="&strTxt&"&selC="&sCategory&"&eventstate="&sState&"&eventkind="&sKind
	'#######################################
	
	'데이터 가져오기
	
	'--이벤트 상품	
	 set cEGroup = new ClsEventGroup
 		cEGroup.FECode = eCode  
 		cEGroup.FEChannel = eChannel 
  		arrGroup = cEGroup.fnGetEventItemGroup	  
 	set cEGroup = nothing

 	if itemsort = "" then itemsort = eisort
	set cEvtItem = new ClsEvent	
		
		cEvtItem.FPSize = iPageSize	
		cEvtItem.FECode = eCode	
		cEvtItem.FRectMakerid = makerid
		cEvtItem.FRectItemid = itemid
		cEvtItem.FRectItemName = itemname 
		cEvtItem.FRectOnlyMobile = blnOnlyMobile
       	cEvtItem.FRectDispCate = dispCate
        cEvtItem.FCPage = iCurrpage
        cEvtItem.FESGroup = strG	
		cEvtItem.FESSort = itemsort	
        cEvtItem.FEChannel = eChannel
 		arrList = cEvtItem.fnGetEventItem 		
 		iTotCnt = cEvtItem.FTotCnt	'전체 데이터  수
        iDispYCnt = cEvtItem.FDispYCnt
        iDispNCnt = cEvtItem.FDispNCnt 
 	    iTotalPage 	=  int((iTotCnt-1)/iPageSize) +1  '전체 페이지 수
%>
<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
<head>
<meta http-equiv="Content-Type" content="text/html;charset=utf-8" />
<style type="text/css">
br { mso-data-placement:same-cell; }
</style>
</head>
<body>
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
        <td>상품ID</td>
        <td>브랜드</td>
        <td>상품명</td>
        <td>판매가</td>
        <td>매입가</td>
        <td>할인율</td>
        <td>배송</td>
        <td>판매여부</td>	
        <td>상품사용여부</td>	
        <td>한정여부</td> 
    </tr>
<%IF isArray(arrList) THEN 
    For intLoop = 0 To UBound(arrList,2)
%>
    <tr align="center" bgcolor="#FFFFFF">
        <td><%=arrList(0,intLoop)%></td>
        <td><%=db2html(arrList(3,intLoop))%></td>
        <td align="left">&nbsp;<%=db2html(arrList(4,intLoop))%></td>
        <td>
            <%
                Response.Write FormatNumber(arrList(7,intLoop),0)
                '할인가
                if arrList(18,intLoop)="Y" then
                    Response.Write "<br><font color=#F08050>(할)" & FormatNumber(arrList(9,intLoop),0) & "</font>"
                end if
                '쿠폰가
                if arrList(22,intLoop)="Y" then
                    Select Case arrList(23,intLoop)
                        Case "1"
                            Response.Write "<br><font color=#5080F0>(쿠)" & FormatNumber(arrList(5,intLoop)*((100-arrList(24,intLoop))/100),0) & "</font>"
                        Case "2"
                            Response.Write "<br><font color=#5080F0>(쿠)" & FormatNumber(arrList(5,intLoop)-arrList(24,intLoop),0) & "</font>"
                    end Select
                end if
            %>
        </td>
        <td>
            <%
                Response.Write FormatNumber(arrList(8,intLoop),0)
                '할인가
                if arrList(18,intLoop)="Y" then
                Response.Write "<br><font color=#F08050>" & FormatNumber(arrList(10,intLoop),0) & "</font>"
                end if
                '쿠폰가
                if arrList(22,intLoop)="Y" then
                if arrList(23,intLoop)="1" or arrList(23,intLoop)="2" then
                    if arrList(25,intLoop)=0 or isNull(arrList(25,intLoop)) then
                        Response.Write "<br><font color=#5080F0>" & FormatNumber(arrList(8,intLoop),0) & "</font>"
                    else
                        Response.Write "<br><font color=#5080F0>" & FormatNumber(arrList(25,intLoop),0) & "</font>"
                    end if
                end if
                end if
            %>
        </td>
        <td>
            <%if arrList(18,intLoop)="Y" then%>
            <font color=#F08050><%=CLng(((arrList(7,intLoop)-arrList(9,intLoop))/arrList(7,intLoop))*100)%>%</font>		
            <%end if%>
            <%
                if arrList(22,intLoop)="Y" then 
                    if arrList(23,intLoop)="1" or arrList(23,intLoop)="2" then
                        if arrList(25,intLoop)=0 or isNull(arrList(25,intLoop)) then
                                Response.Write "<br><font color=#5080F0>" & FormatNumber( arrList(8,intLoop),0) & "</font>"
                        else
                            Response.Write "<br><font color=#5080F0>" & FormatNumber(arrList(24,intLoop),0) 
                                if arrList(23,intLoop)="1" then 
                                Response.Write "%"
                            else
                                Response.Write "원"
                            end if
                                Response.Write "</font>"
                        end if
                    end if
                end if
            %>
        </td>
        <td><%= fnColor(cEvtItem.IsUpcheBeasong(arrList(15,intLoop)),"delivery")%></td>
        <td><%= fnColor(arrList(14,intLoop),"yn") %></td>
        <td><%= fnColor(arrList(19,intLoop),"yn") %></td>
        <td><%= fnColor(arrList(16,intLoop),"yn") %></td>
    </tr>
<% Next %>
<% end if %>
</table>
</body>
</html>
<% session.codePage = 949 %>
<!-- #include virtual="/lib/db/dbclose.asp" -->