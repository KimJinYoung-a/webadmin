<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/designer_cs_baljucls.asp"-->
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

set ojumun = new CCSJumunMaster

ojumun.FRectOrderSerial = listitem
ojumun.FRectIsAll       = iSall
ojumun.FRectDesignerID = session("ssBctID")
ojumun.reDesignerCS_SelectBaljuList

dim IsMeaipPriceValid : IsMeaipPriceValid = session("ssBctID")="esopoom"
dim i, j


'Response.Buffer=False
Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=TEN" & Left(CStr(now()),10) & "_" & session.sessionID & ".csv"
Response.CacheControl = "public"

dim bufStr, tmpS
bufStr = ""

if (IsMeaipPriceValid) then
    bufStr = "��������,�Ϸù�ȣ,CS�ֹ���ȣ,������,�����ڸ�,��������ȭ,�������ڵ���,������,��������ȭ,�������ڵ���,�����ȣ,������ּ�1,������ּ�2,��ǰ�ڵ�,��ǰ��,�ɼ�,�ǸŰ�,����,��ü��ǰ�ڵ�,���԰�"
else
    bufStr = "��������,�Ϸù�ȣ,CS�ֹ���ȣ,������,�����ڸ�,��������ȭ,�������ڵ���,������,��������ȭ,�������ڵ���,�����ȣ,������ּ�1,������ּ�2,��ǰ�ڵ�,��ǰ��,�ɼ�,�ǸŰ�,����,��ü��ǰ�ڵ�"
end if

response.write bufStr & VbCrlf

for ix=0 to ojumun.FResultCount - 1
    requiredetailArr = ""
    bufStr = ""
    bufStr = bufStr & Chr(34) & ojumun.FMasterItemList(ix).Fdivcdname & Chr(34)
    bufStr = bufStr & "," & Chr(34) & ojumun.FMasterItemList(ix).Fcsdetailidx & Chr(34)
    bufStr = bufStr & "," & Chr(34) & ojumun.FMasterItemList(ix).FOrgOrderSerial & "-" & ojumun.FMasterItemList(ix).Fcsmasteridx & Chr(34)
    bufStr = bufStr & "," & Chr(34) & Left(CStr(ojumun.FMasterItemList(ix).FRegDate),10) & Chr(34)
    bufStr = bufStr & "," & Chr(34) & ReplaceSCVStr(ojumun.FMasterItemList(ix).FBuyName) & Chr(34)
    bufStr = bufStr & "," & Chr(34) & ReplaceSCVStr(ojumun.FMasterItemList(ix).FBuyPhone) & Chr(34)
    bufStr = bufStr & "," & Chr(34) & ReplaceSCVStr(ojumun.FMasterItemList(ix).FBuyHp) & Chr(34)
    bufStr = bufStr & "," & Chr(34) & ReplaceSCVStr(ojumun.FMasterItemList(ix).FReqName) & Chr(34)
    bufStr = bufStr & "," & Chr(34) & ReplaceSCVStr(ojumun.FMasterItemList(ix).FReqPhone) & Chr(34)
    bufStr = bufStr & "," & Chr(34) & ReplaceSCVStr(ojumun.FMasterItemList(ix).FReqHp) & Chr(34)
    bufStr = bufStr & "," & Chr(34) & ReplaceSCVStr(ojumun.FMasterItemList(ix).FReqZipCode) & Chr(34)
    bufStr = bufStr & "," & Chr(34) & ReplaceSCVStr(ojumun.FMasterItemList(ix).FReqZipAddr) & Chr(34)
    bufStr = bufStr & "," & Chr(34) & ReplaceSCVStr(ojumun.FMasterItemList(ix).FReqAddress) & Chr(34)
    bufStr = bufStr & "," & Chr(34) & ReplaceSCVStr(ojumun.FMasterItemList(ix).Fitemid) & Chr(34)
    bufStr = bufStr & "," & Chr(34) & ReplaceSCVStr(ojumun.FMasterItemList(ix).FItemName) & Chr(34)
    bufStr = bufStr & "," & Chr(34) & ReplaceSCVStr(ojumun.FMasterItemList(ix).FItemoptionName) & Chr(34)
    bufStr = bufStr & "," & Chr(34) & ojumun.FMasterItemList(ix).FItemCost & Chr(34)
    bufStr = bufStr & "," & Chr(34) & ojumun.FMasterItemList(ix).FItemNo & Chr(34)
    bufStr = bufStr & "," & Chr(34) & ReplaceSCVStr(ojumun.FMasterItemList(ix).FupcheManageCode) & Chr(34)

    if (IsMeaipPriceValid) then
        ''���԰� 2011-01 �߰�, ��ۺ� ���� �߰����
        bufStr = bufStr & "," & Chr(34) & ojumun.FMasterItemList(ix).FBuycash &  Chr(34)
        '''bufStr = bufStr & "," & Chr(34) & Chr(34)
    end if
    response.write bufStr & VbCrlf
next %>
<%
set ojumun = Nothing
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
