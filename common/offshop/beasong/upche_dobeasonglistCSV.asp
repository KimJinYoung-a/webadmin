<%@ language=vbscript %>
<%
option explicit
Response.Expires = -1
%>
<%
'###########################################################
' Description : �������� ���
' Hieditor : 2011.02.26 �ѿ�� ����
'###########################################################
%>
<!-- #include virtual="/common/incSessionAdminOrUpche.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/upche/upchebeasong_cls.asp" -->

<%
dim ojumun ,ix,sql ,detailidxarr ,iSall, SheetType ,i, j
dim bufStr, tmpS
	detailidxarr =  request("detailidxarr")
	iSall   =  request("isall")
	SheetType  =  request("SheetType")

	bufStr = ""

If session("ssBctId") = "" then
    response.write "<script language='javascript'>alert('������ ����Ǿ����ϴ�.');</script>"
    dbget.close()	:	response.End
end if

function ReplaceSCVStr(oStr)
    ReplaceSCVStr = ""
    if IsNULL(oStr) then Exit function
    ReplaceSCVStr = Replace(oStr, chr(34),"'")

end function

set ojumun = new cupchebeasong_list
	ojumun.FRectdetailidxarr = detailidxarr
	ojumun.FRectIsAll       = iSall
	ojumun.FRectDesignerID = session("ssBctID")
	ojumun.fReDesignerSelectBaljuList()

response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=TEN" & Left(CStr(now()),10) & "_" & session.sessionID & ".csv"
Response.CacheControl = "public"

bufStr = "�ֹ���ȣ,�ֹ���,�����ڸ�,��������ȭ,�������ڵ���,������,��������ȭ,�������ڵ���,�����ȣ,������ּ�1,������ּ�2,������ǻ���,�ù��ȣ,��ǰ���̵�,��ǰ��,�ɼ�,�ǸŰ�,����"

response.write bufStr & VbCrlf

for ix=0 to ojumun.FResultCount - 1
    bufStr = ""
    bufStr = bufStr & Chr(34) & ojumun.FItemList(ix).Forderno & Chr(34)
    bufStr = bufStr & "," & Chr(34) & Left(CStr(ojumun.FItemList(ix).FRegDate),10) & Chr(34)
    bufStr = bufStr & "," & Chr(34) & ReplaceSCVStr(ojumun.FItemList(ix).FBuyName) & Chr(34)
    bufStr = bufStr & "," & Chr(34) & ReplaceSCVStr(ojumun.FItemList(ix).FBuyPhone) & Chr(34)
    bufStr = bufStr & "," & Chr(34) & ReplaceSCVStr(ojumun.FItemList(ix).FBuyHp) & Chr(34)
    bufStr = bufStr & "," & Chr(34) & ReplaceSCVStr(ojumun.FItemList(ix).FReqName) & Chr(34)
    bufStr = bufStr & "," & Chr(34) & ReplaceSCVStr(ojumun.FItemList(ix).FReqPhone) & Chr(34)
    bufStr = bufStr & "," & Chr(34) & ReplaceSCVStr(ojumun.FItemList(ix).FReqHp) & Chr(34)
    bufStr = bufStr & "," & Chr(34) & ReplaceSCVStr(ojumun.FItemList(ix).FReqZipCode) & Chr(34)
    bufStr = bufStr & "," & Chr(34) & ReplaceSCVStr(ojumun.FItemList(ix).FReqZipAddr) & Chr(34)
    bufStr = bufStr & "," & Chr(34) & ReplaceSCVStr(ojumun.FItemList(ix).FReqAddress) & Chr(34)
    bufStr = bufStr & "," & Chr(34) & ReplaceSCVStr(db2html(ojumun.FItemList(ix).FComment)) & Chr(34)
    bufStr = bufStr & "," & Chr(34) & ReplaceSCVStr(ojumun.FItemList(ix).Fsongjangno) & Chr(34)
    bufStr = bufStr & "," & Chr(34) & ReplaceSCVStr(ojumun.FItemList(ix).Fitemid) & Chr(34)
    bufStr = bufStr & "," & Chr(34) & ReplaceSCVStr(ojumun.FItemList(ix).FItemName) & Chr(34)
    bufStr = bufStr & "," & Chr(34) & ReplaceSCVStr(ojumun.FItemList(ix).FItemoptionName) & Chr(34)
    bufStr = bufStr & "," & Chr(34) & ojumun.FItemList(ix).Fsellprice & Chr(34)
    bufStr = bufStr & "," & Chr(34) & ojumun.FItemList(ix).FItemNo & Chr(34)

    response.write bufStr & VbCrlf
next %>
<%
set ojumun = Nothing
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->