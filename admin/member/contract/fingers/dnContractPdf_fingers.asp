<%@ language=vbscript %><% option explicit %>
<!-- #include virtual="/lib/util/md5.asp"-->
<%

dim i
dim agreeIdx, gkey, ekey, ctrNo, ctrName, vTp, pTp, chkcf, cTp

agreeIdx  = request("agreeIdx")
gkey    = request("gkey")
ekey    = request("ekey")
ctrNo   = request("ctrNo")
cTp     = request("cTp")    
vTp     = request("vTp")
pTp     = request("pTp")
chkcf   = request("chkcf")


''
if cTp="16" then
    ctrName = "�Ǹ����̿���"
elseif cTp="17" then
    ctrName = "�ŷ���༭"    
else
    ctrName = "��༭"
end if


if (ekey="") then
    response.write "��ȣȭ Ű�� �ùٸ��� �ʽ��ϴ�."
    response.end
end if

if (UCASE(ekey)<>UCASE(MD5("TBTCTR"&agreeIdx&gkey))) then
    response.write "��ȣȭ Ű�� �ùٸ��� �ʽ��ϴ�."
    response.end
end if

dim pdfDoc, theData, pdfID
Set pdfDoc = Server.CreateObject("ABCpdf9.Doc")
pdfDoc.SetInfo 0, "HostWebBrowser", "0"
pdfDoc.Color.String = "255 255 255"

pdfDoc.Rect.String = "55 55 556 715"
pdfDoc.HPos = 0.5
pdfDoc.VPos = 0.5

pdfID = pdfDoc.AddImageUrl("http://testwebadmin.10x10.co.kr/lectureadmin/contract/ifrconfirmContract.asp?agreeIdx="&agreeIdx&"&gkey="&gkey&"&ekey="&ekey&"&chkcf="&chkcf)
Do
	'// ���� ������
	pdfDoc.FrameRect
	If Not pdfDoc.Chainable(pdfID) Then Exit Do
	pdfDoc.Page = pdfDoc.AddPage()
	pdfID = pdfDoc.AddImageToChain(pdfID)
Loop

'''-- Ǫ���߰�
pdfDoc.Color.String = "0 0 0"
pdfDoc.Rect.String = "40 740 556 755"
pdfDoc.HPos = 1.0
pdfDoc.VPos = 0.5
pdfDoc.FontSize = 9
dim footCNT : footCNT=pdfDoc.PageCount
if (pTp="1")  then
    footCNT = footCNT-1  '' ������ ������ (�������� ���� �� ����)
end if

For i = 1 To footCNT
  pdfDoc.PageNumber = i
  pdfDoc.AddText "No. "&ctrNo
  'pdfDoc.FrameRect ''������
Next


For i = 1 To pdfDoc.PageCount
	pdfDoc.PageNumber = i
	pdfDoc.Flatten
Next


theData = pdfDoc.GetData()

if (vTp="d") then
    Response.ContentType = "application/octet-stream"  ''�ٿ�ε��
else
    Response.ContentType = "application/pdf" ''������ ����
end if

''Response.AddHeader "content-length", UBound(theData) - LBound(theData) + 1
Response.AddHeader "content-disposition", "inline; filename="&ctrName&".pdf"
Response.BinaryWrite theData

pdfDoc.Clear()

Set pdfDoc = Nothing

%>
