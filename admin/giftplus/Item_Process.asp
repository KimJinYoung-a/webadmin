<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  기프트플러스
' History : 2010.04.05 한용민 생성
'###########################################################
%>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/giftplus/giftplus_cls.asp"-->
<%
dim mode , cdL, cdM, cdS ,div,varTABLE,listType ,arrItemID, arrOrderNo ,ecdL,ecdM , ecdS
dim strSQL,msg
	mode=request("mode")
	cdL = request("cdL")
	cdM = request("cdM")
	cdS = request("cdS")
	arrItemID = chkarray(request("arrItemID"))
	arrOrderNo = chkarray(request("arrOrderNo"))
	ecdL  = request("ecdL")
	ecdM  = request("ecdM")
	ecdS  = request("ecdS")	

	varTABLE = " [db_giftplus].[dbo].[tbl_giftplus_item] "

Public Function getWhereSQL(byval cL,byval cM, byval cS ,byval aITEM)
Dim tSQL

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
'// 상품 삭제
	strSQL =" DELETE From " & varTABLE & " " &_
			" WHERE 1=1 " & getWhereSQL(cdL,cdM,cdS,arritemID)
	
	'response.write strSQL
	msg="OK"
ELSEIF mode= "move" THEN
'// 상품 이동?
	strSQL =" UPDATE " & varTABLE & " " &_
			" set LCode='9999',MCode='9999' , SCode='9999'"&_
			" WHERE 1=1 " & getWhereSQL(cdL,cdM,cdS,arritemID)

	strSQL =strSQL & _
			" UPDATE " & varTABLE & " " &_
			" set LCode='" & ecdL & "',MCode='" & ecdM & "' , SCode='" & ecdS & "'"&_
			" where LCode='9999' and MCode='9999' and SCode='9999'" &_
			"	and itemid not in ( " &_
			" 	SELECT itemid FROM " & varTABLE & "  " &_
			" WHERE 1=1 " & getWhereSQL(ecdL,ecdM,ecdS,arritemID) &_
			") "
	strSQL =strSQL & _
			" DELETE " & varTABLE & " " &_
			" where LCode='9999' and MCode='9999' and SCode='9999'"

	'response.write strSQL
	msg="OK"
ELSEIF mode="copy" THEN
'// 상품 복사 ?
	strSQL =" INSERT INTO " & varTABLE & " (LCode,MCode,SCode,ItemID) " &_
			" SELECT '" & ecdL & "','" & ecdM & "','" & ecdS & "',itemid  " &_
			" FROM " & varTABLE & "  " &_
			" WHERE 1=1 " & getWhereSQL(cdL,cdM,cdS,arritemID) &_
			" and itemid not in  ( " &_
			"	SELECT itemid  " &_
			"	FROM " & varTABLE & "  " &_
			" WHERE 1=1 " & getWhereSQL(ecdL,ecdM,ecdS,arritemID) &_
			") "
	msg="OK"
'response.write strSQL
ELSEIF mode="update" THEN
	
	dim icnt,vcnt,i
	arrItemID	= split(arrItemID,",")
	arrOrderNo  = split(arrOrderNo,",")
	
	icnt = Ubound(arrItemID)
	vcnt = Ubound(arrOrderNo)
	'response.write icnt &"/" &vcnt

	if (icnt>0 and vcnt>0) and (icnt = vcnt) then

		for i=0 to icnt

			if trim(arrOrderNo(i))="" then arrOrderNo(i)="99"

			strSQL = strSQL &_
					" UPDATE " & varTABLE & " " &_
					" SET OrderNo = '" & arrOrderNo(i) & "'" &_
					" WHERE 1=1 " & getWhereSQL(cdL,cdM,cdS,arrItemID(i)) & ";"
		next
	else

		response.write	"<script language='javascript'>" &_
					"	alert('처리중 에러가 발생했습니다.');" &_
					"</script>"
		dbget.close()	:	response.End
	end if

	'response.write strSQL
	msg ="OK"

ELSE
'// 상품 추가
	strSQL =" INSERT INTO " & varTABLE & " (LCode,MCode,SCode,ItemID)" &_
			" SELECT '" & cdL & "','" & cdM & "','" & cdS & "' ,itemid" &_
			" FROM [db_item].[dbo].[tbl_item] " &_
			" WHERE itemid in (" & arrItemID &")" &_
			" and itemid not in ( "&_
			"	SELECT itemid  "&_
			"	FROM " & varTABLE & "  "&_
			" WHERE 1=1 " & getWhereSQL(cdL,cdM,cdS,arrItemID) &_
			"	)"
	msg="OK"
	'response.write strsQL
END IF

	dbget.BeginTrans

	'response.write strsQL
	dbget.execute(strSQL)
	'dbget.close()	:	response.End
	
	'오류검사 및 반영
	If Err.Number = 0 Then
		dbget.CommitTrans				'커밋(정상)

		response.write	"<script language='javascript'>"
		response.write	"	opener.location.replace('iframe_itemList.asp?cdL=" & cdL & "&cdM=" & cdM & "&cdS=" & cdS & "'); self.focus(); alert('" & msg & "'); self.close();"
		response.write	"</script>"
		dbget.close()	:	response.End
	Else
		dbget.RollBackTrans				'롤백(에러발생시)

		response.write	"<script language='javascript'>" &_
					"	alert('처리중 에러가 발생했습니다.');" &_
					"	history.go(-1);" &_
					"</script>"
		dbget.close()	:	response.End
	End If
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->