<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbAcademyopen.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lectureadmin/lib/classes/jumun/baljucls.asp"-->
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

listitem =  request("orderserial")
iSall   =  RequestCheckvar(request("isall"),10)
SheetType  =  RequestCheckvar(request("SheetType"),10)
  	if listitem <> "" then
		if checkNotValidHTML(listitem) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('��ȿ���� ���� ���ڰ� ���ԵǾ� �ֽ��ϴ�. �ٽ� �ۼ� ���ּ���');"
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

bufStr = "�ֹ���ȣ,�ֹ���,�����ڸ�,��������ȭ,�������ڵ���,�������̸���,������,��������ȭ,�������ڵ���,�����ȣ,������ּ�1,������ּ�2,������ǻ���,�ù��ȣ,��ǰ���̵�,��ǰ��,�ɼ�,�ǸŰ�,����,�ֹ����۸޼���,��ü��ǰ�ڵ�,����ǰ,��������,ī�帮��,�޼���,�����»��"
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
                requiredetailArr = requiredetailArr + "[" & (i+1) & "�� ��ǰ ����]" &VbCrLF& splitValue(ojumun.FMasterItemList(ix).Frequiredetail,CAddDetailSpliter,i)&VbCrLF
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
        bufStr = bufStr & "," & Chr(34) & Left(CStr(ojumun.FMasterItemList(ix).Freqdate),10) & "�� " & ojumun.FMasterItemList(ix).Freqtime & "��" & Chr(34)
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