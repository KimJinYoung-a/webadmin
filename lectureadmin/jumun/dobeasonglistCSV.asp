<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lectureadmin/lib/classes/jumun/baljucls.asp"-->
<%
If (session("ssBctId") = "") or (session("ssBctDiv") <> "9999" and session("ssBctDiv") > "9") then
    response.write "<script language='javascript'>alert('세션이 종료되었습니다.');</script>"
    dbget.close()	:	response.End
end if

function ReplaceSCVStr(oStr)
    ReplaceSCVStr = ""
    if IsNULL(oStr) then Exit function
    ReplaceSCVStr = Replace(oStr, chr(34),"'")

end function

dim requiredetailArr : requiredetailArr =""
dim ojumun
dim ix,sql
Dim listitemlist,listitem,listitemcount
dim iSall, SheetType

listitem =  request("orderserial")
iSall   =  RequestCheckvar(request("isall"),10)
SheetType  =  RequestCheckvar(request("SheetType"),10)
  	if listitem <> "" then
		if checkNotValidHTML(listitem) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
		response.write "</script>"
		response.End
		end if
	end if
set ojumun = new CJumunMaster

ojumun.FRectOrderSerial = listitem
ojumun.FRectIsAll       = iSall
ojumun.FRectDesignerID = session("ssBctID")
ojumun.ReDesignerSelectBaljuList

dim i, j

'Response.Buffer=False
Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=TEN" & Left(CStr(now()),10) & "_" & session.sessionID & ".csv"
Response.CacheControl = "public"

dim bufStr, tmpS
bufStr = ""

bufStr = "주문번호,주문일,구매자명,구매자전화,구매자핸드폰,구매자이메일,수령인,수령인전화,수령인핸드폰,우편번호,배송지주소1,배송지주소2,배송유의사항,택배번호,상품아이디,상품명,옵션,판매가,수량,주문제작메세지,업체상품코드,사은품,배송희망일,카드리본,메세지,보내는사람"
response.write bufStr & VbCrlf

for ix=0 to ojumun.FResultCount - 1
    requiredetailArr = ""
    bufStr = ""
    bufStr = bufStr & Chr(34) & ojumun.FMasterItemList(ix).FOrderSerial & Chr(34)
    bufStr = bufStr & "," & Chr(34) & Left(CStr(ojumun.FMasterItemList(ix).FRegDate),10) & Chr(34)
    bufStr = bufStr & "," & Chr(34) & ReplaceSCVStr(ojumun.FMasterItemList(ix).FBuyName) & Chr(34)
    bufStr = bufStr & "," & Chr(34) & ReplaceSCVStr(ojumun.FMasterItemList(ix).FBuyPhone) & Chr(34)
    bufStr = bufStr & "," & Chr(34) & ReplaceSCVStr(ojumun.FMasterItemList(ix).FBuyHp) & Chr(34)
    bufStr = bufStr & "," & Chr(34) & Chr(34)
    bufStr = bufStr & "," & Chr(34) & ReplaceSCVStr(ojumun.FMasterItemList(ix).FReqName) & Chr(34)
    bufStr = bufStr & "," & Chr(34) & ReplaceSCVStr(ojumun.FMasterItemList(ix).FReqPhone) & Chr(34)
    bufStr = bufStr & "," & Chr(34) & ReplaceSCVStr(ojumun.FMasterItemList(ix).FReqHp) & Chr(34)
    bufStr = bufStr & "," & Chr(34) & ReplaceSCVStr(ojumun.FMasterItemList(ix).FReqZipCode) & Chr(34)
    bufStr = bufStr & "," & Chr(34) & ReplaceSCVStr(ojumun.FMasterItemList(ix).FReqZipAddr) & Chr(34)
    bufStr = bufStr & "," & Chr(34) & ReplaceSCVStr(ojumun.FMasterItemList(ix).FReqAddress) & Chr(34)
    bufStr = bufStr & "," & Chr(34) & ReplaceSCVStr(db2html(ojumun.FMasterItemList(ix).FComment)) & Chr(34)
    bufStr = bufStr & "," & Chr(34) & ReplaceSCVStr(ojumun.FMasterItemList(ix).Fsongjangno) & Chr(34)
    bufStr = bufStr & "," & Chr(34) & ReplaceSCVStr(ojumun.FMasterItemList(ix).Fitemid) & Chr(34)
    bufStr = bufStr & "," & Chr(34) & ReplaceSCVStr(ojumun.FMasterItemList(ix).FItemName) & Chr(34)
    bufStr = bufStr & "," & Chr(34) & ReplaceSCVStr(ojumun.FMasterItemList(ix).FItemoptionName) & Chr(34)
    bufStr = bufStr & "," & Chr(34) & ojumun.FMasterItemList(ix).FItemCost & Chr(34)
    bufStr = bufStr & "," & Chr(34) & ojumun.FMasterItemList(ix).FItemNo & Chr(34)
    requiredetailArr=""
    if (ojumun.FMasterItemList(ix).FItemNo>1) then
        if (Not IsNULL(ojumun.FMasterItemList(ix).Frequiredetail)) then
            if (ojumun.FMasterItemList(ix).Frequiredetail<>"") then
            for i=0 to ojumun.FMasterItemList(ix).FItemNo-1
                requiredetailArr = requiredetailArr + "[" & (i+1) & "번 상품 문구]" &VbCrLF& splitValue(ojumun.FMasterItemList(ix).Frequiredetail,CAddDetailSpliter,i)&VbCrLF
            next
            end if
        end if
    bufStr = bufStr & "," & Chr(34) & ReplaceSCVStr(requiredetailArr) & Chr(34)
    else
    bufStr = bufStr & "," & Chr(34) & ReplaceSCVStr(ojumun.FMasterItemList(ix).Frequiredetail) & Chr(34)
    end if
    bufStr = bufStr & "," & Chr(34) & ReplaceSCVStr(ojumun.FMasterItemList(ix).FupcheManageCode) & Chr(34)

    bufStr = bufStr & "," & Chr(34) & Chr(34)

    if Not IsNULL(ojumun.FMasterItemList(ix).Freqdate) then
        bufStr = bufStr & "," & Chr(34) & Left(CStr(ojumun.FMasterItemList(ix).Freqdate),10) & "일 " & ojumun.FMasterItemList(ix).Freqtime & "시" & Chr(34)
        bufStr = bufStr & "," & Chr(34) & ReplaceSCVStr(ojumun.FMasterItemList(ix).getCardribbonName) & Chr(34)
        bufStr = bufStr & "," & Chr(34) & ReplaceSCVStr(db2html(ojumun.FMasterItemList(ix).Fmessage)) & Chr(34)
        bufStr = bufStr & "," & Chr(34) & ReplaceSCVStr(db2html(ojumun.FMasterItemList(ix).Ffromname)) & Chr(34)
    else
        bufStr = bufStr & "," & Chr(34) & Chr(34)
        bufStr = bufStr & "," & Chr(34) & Chr(34)
        bufStr = bufStr & "," & Chr(34) & Chr(34)
        bufStr = bufStr & "," & Chr(34) & Chr(34)
    end if
    response.write bufStr & VbCrlf
next %>
<%
set ojumun = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->