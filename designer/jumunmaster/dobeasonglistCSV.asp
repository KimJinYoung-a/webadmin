<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesignerNoCache.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/designer/checkPartnerLog.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/designer_baljucls.asp"-->
<!-- #include virtual="/lib/classes/order/ordergiftcls.asp"-->
<%
If (session("ssBctId") = "") or (session("ssBctDiv") <> "9999" and session("ssBctDiv") > "9") then
    response.write "<script language='javascript'>alert('������ ����Ǿ����ϴ�.');</script>"
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

listitem =  Replace(request("orderserial"), " ", "")
iSall   =  requestCheckVar(request("isall"), 32)
SheetType  =  requestCheckVar(request("SheetType"), 32)

set ojumun = new CJumunMaster

ojumun.FRectOrderSerial = listitem
ojumun.FRectIsAll       = iSall
ojumun.FRectDesignerID = session("ssBctID")
ojumun.ReDesignerSelectBaljuList

dim IsMeaipPriceValid : IsMeaipPriceValid = session("ssBctID")="esopoom"
dim oGift, i, j

set oGift = new COrderGift

''Response.Buffer=False
Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=TEN" & Left(CStr(now()),10) & "_" & session.sessionID & ".csv"
Response.CacheControl = "public"

dim bufStr, tmpS
bufStr = ""

if (IsMeaipPriceValid) then
    bufStr = "�ֹ���ȣ,�ֹ���,�����ڸ�,��������ȭ,�������ڵ���,�������̸���,������,��������ȭ,�������ڵ���,�����ȣ,������ּ�1,������ּ�2,������ǻ���,�ù��ȣ,��ǰ���̵�,��ǰ��,�ɼ�,�ǸŰ�,����,�ֹ����۸޼���,��ü��ǰ�ڵ�,����ǰ,��������,ī�帮��,�޼���,�����»��,���԰�"
else
    bufStr = "�ֹ���ȣ,�ֹ���,�����ڸ�,��������ȭ,�������ڵ���,�������̸���,������,��������ȭ,�������ڵ���,�����ȣ,������ּ�1,������ּ�2,������ǻ���,�ù��ȣ,��ǰ���̵�,��ǰ��,�ɼ�,�ǸŰ�,����,�ֹ����۸޼���,��ü��ǰ�ڵ�,����ǰ,��������,ī�帮��,�޼���,�����»��"
end if

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
					requiredetailArr = requiredetailArr + "[" & (i+1) & "�� ��ǰ ����]" &" "& splitValue(ojumun.FMasterItemList(ix).Frequiredetail,CAddDetailSpliter,i)&" "
				next
            end if
        end if
		bufStr = bufStr & "," & Chr(34) & ReplaceSCVStr(requiredetailArr) & Chr(34)
    else
		bufStr = bufStr & "," & Chr(34) & ReplaceSCVStr(Replace(ojumun.FMasterItemList(ix).Frequiredetail, CAddDetailSpliter, "")) & Chr(34)
    end if
    bufStr = bufStr & "," & Chr(34) & ReplaceSCVStr(ojumun.FMasterItemList(ix).FupcheManageCode) & Chr(34)

    oGift.FRectOrderSerial = ojumun.FMasterItemList(ix).FOrderSerial
    oGift.FRectMakerid = session("ssBctId")
    oGift.FRectGiftDelivery = "Y"
    oGift.GetOneOrderGiftlist
    if (oGift.FResultCount>0) then
        for j=0 to oGift.FResultCount -1
			tmpS = tmpS & ReplaceSCVStr(oGift.FItemList(j).GetEventConditionStr) & " "
        next
        bufStr = bufStr & "," & Chr(34) & tmpS & Chr(34)
    else
        bufStr = bufStr & "," & Chr(34) & Chr(34)
    end if

    if Not IsNULL(ojumun.FMasterItemList(ix).Freqdate) then
        bufStr = bufStr & "," & Chr(34) & Left(CStr(ojumun.FMasterItemList(ix).Freqdate),10) & "�� " & ojumun.FMasterItemList(ix).GetReqTimeText & Chr(34)
        bufStr = bufStr & "," & Chr(34) & ReplaceSCVStr(ojumun.FMasterItemList(ix).getCardribbonName) & Chr(34)
        bufStr = bufStr & "," & Chr(34) & ReplaceSCVStr(db2html(ojumun.FMasterItemList(ix).Fmessage)) & Chr(34)
        bufStr = bufStr & "," & Chr(34) & ReplaceSCVStr(db2html(ojumun.FMasterItemList(ix).Ffromname)) & Chr(34)
    else
        bufStr = bufStr & "," & Chr(34) & Chr(34)
        bufStr = bufStr & "," & Chr(34) & Chr(34)
        bufStr = bufStr & "," & Chr(34) & Chr(34)
        bufStr = bufStr & "," & Chr(34) & Chr(34)
    end if

    if (IsMeaipPriceValid) then
        ''���԰� 2011-01 �߰�, ��ۺ� ���� �߰����
        bufStr = bufStr & "," & Chr(34) & ojumun.FMasterItemList(ix).FBuycash &  Chr(34)
        '''bufStr = bufStr & "," & Chr(34) & Chr(34)
    end if
    response.write bufStr & VbCrlf
next %>
<%
set ojumun = Nothing
set oGift = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
