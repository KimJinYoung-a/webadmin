<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/designer/incSessionDesigner.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/items/itemcls_v2.asp"-->
<%

dim itemid, makerid, itemname
dim sellyn, isusing, danjongyn, limityn, mwdiv
dim page, arrlist ,bufStr, cdl, cdm, cds
dim infodivYn

itemid  = RequestCheckVar(request("itemid"),10)
makerid = RequestCheckVar(request("makerid"),32)
itemname = RequestCheckVar(request("itemname"),32)

sellyn  = RequestCheckVar(request("sellyn"),10)
isusing = RequestCheckVar(request("isusing"),10)
danjongyn = RequestCheckVar(request("danjongyn"),10)
limityn = RequestCheckVar(request("limityn"),10)
mwdiv = RequestCheckVar(request("mwdiv"),10)

page = RequestCheckVar(request("page"),10)
cdl = requestCheckvar(request("cdl"),10)
cdm = requestCheckvar(request("cdm"),10)
cds = requestCheckvar(request("cds"),10)
infodivYn  = requestCheckvar(request("infodivYn"),10)

if (sellyn="") then sellyn="A"

if (page="") then page=1

''if (isusing="") then isusing="Y"
''����ϴ� ��ǰ�� ǥ�÷� ����
isusing="Y"

'��ǰ�ڵ� ��ȿ�� �˻�(2008.08.01;������)
if itemid<>"" then
	if Not(isNumeric(itemid)) then
		Response.Write "<script language=javascript>alert('[" & itemid & "]��(��) ��ȿ�� ��ǰ�ڵ尡 �ƴմϴ�.');history.back();</script>"
		dbget.close()	:	response.End
	end if
end if

'==============================================================================
dim oitem

set oitem = new CItem

oitem.FRectMakerId = session("ssBctID")
oitem.FRectItemid = itemid
oitem.FRectItemName = itemname
oitem.FRectDanjongyn = danjongyn
oitem.FRectLimityn = limityn
oitem.FRectMWDiv = mwdiv
oitem.FPageSize = 30
oitem.FRectIsExcelDown = "o"
oitem.FCurrPage = page
oitem.FRectCate_Large   = cdl
oitem.FRectCate_Mid     = cdm
oitem.FRectCate_Small   = cds
oitem.FRectInfodivYn    = infodivYn

if (sellyn <> "A") then
    oitem.FRectSellYN = sellyn
end if

if (isusing <> "A") then
    oitem.FRectIsUsing = isusing
end if

if (oitem.FRectMakerId<>"") then
    oitem.GetProductListcsv
end if

'//�����ҽ� ���� �迭�� �޾ƿ�. �޸� ��� ������		'/2012.11.06 �ѿ�� �߰�
if oitem.FresultCount > 0 then
	arrlist = oitem.farrlist
end if

dim i

'Response.Buffer=False
Response.Expires=0
response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment; filename=��ǰ����Ʈ.csv"
Response.CacheControl = "public"

response.write "��ǰ�ڵ�,��ǰ��,�ŷ�����,�Ǹſ���,��������,��������,�Һ��ڰ�,�ǸŰ�,���ް�,��ü�ڵ�,��۱���" & VbCrlf

if isarray(arrlist) then
	for i=0 to ubound(arrlist,2)
	    bufStr = "" 
	    bufStr = bufStr & arrlist(0,i)
	    bufStr = bufStr & "," & Chr(34) & arrlist(2,i) & Chr(34)
	    bufStr = bufStr & "," & mwdivName(arrlist(7,i))
	    bufStr = bufStr & "," & arrlist(5,i)
	    bufStr = bufStr & "," & arrlist(8,i)

		if (arrlist(8,i) = "Y") then
			bufStr = bufStr & "," & getLimitEa(arrlist(9,i),arrlist(10,i))
		else
			bufStr = bufStr & ","
		end if

        bufStr = bufStr & "," & arrlist(15,i)
	    bufStr = bufStr & "," & arrlist(3,i)
	    bufStr = bufStr & "," & arrlist(4,i)
	    bufStr = bufStr & "," & arrlist(13,i)
	    bufStr = bufStr & ","
	    
		If arrlist(14,i) = "1" Then
			bufStr = bufStr & "�ٹ����ٹ��"
		ElseIf arrlist(14,i) = "2" Then
			bufStr = bufStr & "��ü(����)���"
		ElseIf arrlist(14,i) = "4" Then
			bufStr = bufStr & "�ٹ����ٹ�����"
		ElseIf arrlist(14,i) = "9" Then
			bufStr = bufStr & "��ü���ǹ��(���� ��ۺ�ΰ�)"
		ElseIf arrlist(14,i) = "7" Then
			bufStr = bufStr & "��ü���ҹ��"
		End If
	    
	    response.write bufStr & VbCrlf
	next
end if
%>

<% set oitem = nothing %>

<!-- #include virtual="/lib/db/dbclose.asp" -->