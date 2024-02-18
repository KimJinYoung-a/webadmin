<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/giftManager/GiftManagerCls.asp"-->>

<%

dim mode

mode=request("mode")

dim cdL,cdM,cdS

cdL = request("cdL")
cdM = request("cdM")
cdS = request("cdS")

dim div,varTABLE
div = request("div")

if div<>"" then
	varTABLE = "[db_giftManager].[dbo].[tbl_gift_BestItem]"
else
	varTABLE = "[db_giftManager].[dbo].[tbl_gift_item]"
end if


dim arrItemID,arrOrderNo

arrItemID = chkarray(request("arrItemID"))
arrOrderNo = chkarray(request("arrOrderNo"))

dim ecdL,ecdM , ecdS

ecdL  = request("ecdL")
ecdM  = request("ecdM")
ecdS  = request("ecdS")

dim strSQL,msg

Public Function getWhereSQL(byval Dv, byval cL,byval cM, byval cS ,byval aITEM)

	Dim tSQL

	If Dv<>"" Then
		tSQL = tSQL & " and DIV='" & Dv & "'"
	END IF

	IF cL<>"" THEN
		tSQL = tSQL & " and LCode ='" & cL & "'"
	END IF

	IF cM<>"" THEN
		tSQL = tSQL & " and MCode ='" & cM & "'"
	end if

	IF cS<>"" THEN
		tSQL = tSQL & " and SCode ='" & cS & "'"
	END IF

	IF aITEM<>"" THEN
		tSQL = tSQL & " and ItemID in (" & aITEM &") "
	END IF
	getWhereSQL = tSQL
End Function

IF mode = "del" THEN
'// ��ǰ ����
	strSQL =" DELETE " & varTABLE & " " &_
			" WHERE 1=1 " & getWhereSQL(div,cdL,cdM,cdS,arritemID)


	msg="���� �Ǿ����ϴ�"
ELSEIF mode= "move" THEN
'// ��ǰ �̵�?
	strSQL =" UPDATE " & varTABLE & " " &_
			" set LCode='9999',MCode='9999' , SCode='9999'"&_
			" WHERE 1=1 " & getWhereSQL(div,cdL,cdM,cdS,arritemID)

	strSQL =strSQL & _
			" UPDATE " & varTABLE & " " &_
			" set LCode='" & ecdL & "',MCode='" & ecdM & "' , SCode='" & ecdS & "'"&_
			" where LCode='9999' and MCode='9999' and SCode='9999'" &_
			"	and itemid not in ( " &_
			" 	SELECT itemid FROM " & varTABLE & "  " &_
			" WHERE 1=1 " & getWhereSQL(div,ecdL,ecdM,ecdS,arritemID) &_
			") "
	strSQL =strSQL & _
			" DELETE " & varTABLE & " " &_
			" where LCode='9999' and MCode='9999' and SCode='9999'"


	msg="��ǰ�� �̵��Ͽ����ϴ�."
ELSEIF mode="copy" THEN
'// ��ǰ ���� ?
	strSQL =" INSERT INTO " & varTABLE & " (LCode,MCode,SCode,ItemID) " &_
			" SELECT '" & ecdL & "','" & ecdM & "','" & ecdS & "',itemid  " &_
			" FROM " & varTABLE & "  " &_
			" WHERE 1=1 " & getWhereSQL(div,cdL,cdM,cdS,arritemID) &_
			" and itemid not in  ( " &_
			"	SELECT itemid  " &_
			"	FROM " & varTABLE & "  " &_
			" WHERE 1=1 " & getWhereSQL(div,ecdL,ecdM,ecdS,arritemID) &_
			") "
	msg="���� �Ǿ����ϴ�?

ELSEIF mode="update" THEN
	dim icnt,vcnt,i

	arrItemID	= split(arrItemID,",")
	arrOrderNo  = split(arrOrderNo,",")

	icnt = Ubound(arrItemID)
	vcnt = Ubound(arrOrderNo)


	if (icnt>0 and vcnt>0) and (icnt = vcnt) then

		for i=0 to icnt

			if trim(arrOrderNo(i))="" then arrOrderNo(i)="99"

			strSQL = strSQL &_
					" UPDATE " & varTABLE & " " &_
					" SET OrderNo = '" & arrOrderNo(i) & "'" &_
					" WHERE 1=1 " & getWhereSQL(div,cdL,cdM,cdS,arrItemID(i)) & ";"
		next
	else

		response.write	"<script language='javascript'>" &_
					"	alert('ó���� ������ �߻��߽��ϴ�.');" &_
					"	history.go(-1);" &_
					"</script>"
		dbget.close()	:	response.End
	end if


	msg ="���� �Ǿ����ϴ�"

ELSE
'// ��ǰ �߰�
	strSQL =" INSERT INTO " & varTABLE & " (LCode,MCode,SCode,ItemID)" &_
			" SELECT '" & cdL & "','" & cdM & "','" & cdS & "' ,itemid" &_
			" FROM [db_item].[dbo].[tbl_item] " &_
			" WHERE itemid in (" & arrItemID &")" &_
			" and itemid not in ( "&_
			"	SELECT itemid  "&_
			"	FROM " & varTABLE & "  "&_
			" WHERE 1=1 " & getWhereSQL(div,cdL,cdM,cdS,arrItemID) &_
			"	)"
	msg="�߰� �Ǿ����ϴ�"
END IF
	dbget.BeginTrans

	response.write strsQL
	dbget.execute(strSQL)

	'�����˻� �� �ݿ�
	If Err.Number = 0 Then
		dbget.CommitTrans				'Ŀ��(����)

		response.write	"<script language='javascript'>"
		response.write	" alert('" & msg & "'); opener.location.replace?'iframe_itemList.asp?cdL=" & cdL & "&cdM=" & cdM & "&cdS=" & cdS & "');self.close();"
		response.write	"</script>"
		dbget.close()	:	response.End
	Else
		dbget.RollBackTrans				'�ѹ�(�����߻���)

		response.write	"<script language='javascript'>" &_
					"	alert('ó���� ������ �߻��߽��ϴ�.');" &_
					"	history.go(-1);" &_
					"</script>"


	End If
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->