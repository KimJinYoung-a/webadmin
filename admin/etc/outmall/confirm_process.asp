<%@ language=vbscript %>
<% option explicit %>
<% Server.ScriptTimeOut = 600 %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbCTopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<%
Dim cmdparam : cmdparam = requestCheckVar(request("cmdparam"),30)
Dim cksel : cksel = request("cksel")
Dim i, iidx, ret, ierrStr
Dim SuccCNT : SuccCNT = 0
Dim FailCNT : FailCNT = 0
Dim sugiMallid, sugisellyn, sugimakerid, sugiadminid, sugiadminText, mallidArr, arrstandardMargin, vMargins
Dim sugiSQL

sugiMallid			= request("sugiMallid")
sugisellyn			= request("sugisellyn")
sugimakerid			= request("sugimakerid")
sugiadminid			= request("sugiadminid")
sugiadminText		= request("sugiadminText")
arrstandardMargin	= request("arrstandardMargin")

Function fnOverseasMall(imallgubun)
	Dim strSQL, cnt
	strSQL = ""
	strSQL = strSQL & " SELECT COUNT(*) as cnt " & VBCRLF
	strSQL = strSQL & " FROM db_user.dbo.tbl_user_c c "
	strSQL = strSQL & " JOIN db_partner.dbo.tbl_partner_addInfo f on c.userid=f.partnerid and f.pcomType=1 and f.pmallSellType=8 "
	strSQL = strSQL & " WHERE c.isusing='Y' and c.userdiv='50' "
	strSQL = strSQL & " and c.userid = '"& imallgubun &"' "
	rsget.CursorLocation = adUseClient
	rsget.Open strSQL, dbget, adOpenForwardOnly, adLockReadOnly
		cnt = rsget("cnt")
	rsget.Close
	If cnt > 0 Then
		fnOverseasMall = "Y"
	Else
		fnOverseasMall = "N"
	End If
End Function

If (cmdparam="confirmOK") Then
	cksel = split(cksel, ",")
	For i = 0 To UBound(cksel)
		iidx = Trim(cksel(i))
		ret = confirmOneItem(iidx, ierrStr)
	Next
	response.write "<script>alert('반영되었습니다.');parent.location.reload();</script>"
ElseIf (cmdparam="marginOK") Then
	cksel		= split(cksel, ",")
	If Right(arrstandardMargin,4) = "*(^!" Then
		arrstandardMargin = Left(arrstandardMargin, Len(arrstandardMargin) - 4)
	End If
	vMargins	= split(arrstandardMargin, "*(^!")

	For i = 0 to UBound(cksel)
		sugiSQL = ""
		sugiSQL = sugiSQL & " UPDATE db_partner.dbo.tbl_partner_addInfo "
		sugiSQL = sugiSQL & " SET outmallstandardMargin = '"& Trim(vMargins(i)) &"' "
		sugiSQL = sugiSQL & " WHERE partnerid = '"& Trim(cksel(i)) &"' "
		dbget.Execute sugiSQL
	Next
	response.write "<script>alert('반영되었습니다.');parent.location.reload();</script>"
ElseIf (cmdparam="sugiOK") Then
	If Trim(sugiadminText) = "" Then
		response.write "<script>alert('사유를 정확히 입력하세요.');opener.location.reload();window.close();</script>"
		response.end
	End If

	If sugiMallid = "all" Then
		sugiSQL = ""
		sugiSQL = sugiSQL & " UPDATE db_etcmall.dbo.tbl_jaehumall_hopeSell SET " & vbcrlf
		sugiSQL = sugiSQL & " isComplete = 'X' " & vbcrlf
		sugiSQL = sugiSQL & " WHERE makerid = '"&sugimakerid&"' and makerid <> 'nvstorefarm' and mallgubun <> 'all' and mallgubun <> 'daumep' and mallgubun <> 'naverep' and mallgubun <> 'shodocep' and mallgubun <> 'wemakepriceep' and mallgubun <> 'ggshop' " & vbcrlf
		dbget.Execute sugiSQL
	End If

	mallidArr = Split(sugiMallid, ",")
	for i = 0 to UBound(mallidArr)
		if (Trim(mallidArr(i)) <> "") then
			sugiSQL = ""
			sugiSQL = sugiSQL & " IF EXISTS(SELECT TOP 1 * FROM db_etcmall.dbo.tbl_jaehumall_hopeSell WHERE makerid='"&sugimakerid&"' and mallgubun='"&Trim(mallidArr(i))&"' and currstat=2 and iscomplete <> 'X' )" & vbcrlf
			sugiSQL = sugiSQL & " 	BEGIN " & vbcrlf
			sugiSQL = sugiSQL & " 		UPDATE db_etcmall.dbo.tbl_jaehumall_hopeSell SET " & vbcrlf
			sugiSQL = sugiSQL & " 		whyhope = '"&html2db(sugiadminText)&"' " & vbcrlf
			sugiSQL = sugiSQL & " 		,currstat=1 " & vbcrlf
			sugiSQL = sugiSQL & " 		,hoperegdate = getdate() " & vbcrlf
			sugiSQL = sugiSQL & " 		WHERE makerid='"&sugimakerid&"' and mallgubun='"&Trim(mallidArr(i))&"' and currstat=2  " & vbcrlf
			sugiSQL = sugiSQL & " 	END " & vbcrlf
			sugiSQL = sugiSQL & " ELSE " & vbcrlf
			sugiSQL = sugiSQL & " 	BEGIN " & vbcrlf
			sugiSQL = sugiSQL & " 		INSERT INTO db_etcmall.dbo.tbl_jaehumall_hopeSell (makerid, mallgubun, currstat, hopesellstat, whyhope, hoperegdate, isComplete) " & vbcrlf
			sugiSQL = sugiSQL & " 		VALUES ('"&sugimakerid&"', '"&Trim(mallidArr(i))&"', '1', '"&sugisellyn&"', '"&html2db(sugiadminText)&"', getdate(), 'N') " & vbcrlf
			sugiSQL = sugiSQL & " 	END " & vbcrlf
			dbget.Execute sugiSQL

			sugiSQL = ""
			sugiSQL = sugiSQL & " INSERT INTO db_etcmall.dbo.tbl_jaehumall_hopeSell_Log (mallgubun, makerid, hopeStr, useYN, reguserid, regdate) " & vbcrlf
			sugiSQL = sugiSQL & " VALUES ('"&Trim(mallidArr(i))&"', '"&sugimakerid&"', '"&html2db(sugiadminText)&"', '"&sugisellyn&"', '"&sugiadminid&"', getdate()) " & vbcrlf
			dbget.Execute sugiSQL
		end if
	next

	response.write "<script>alert('반영되었습니다.');opener.location.reload();window.close();</script>"
End If

Function confirmOneItem(iidx, ierrStr)
	Dim strSQL, fnmallgubun, fnmakerid, fnhopesellstat
	Dim currExtusing
	strSQL = ""
	strSQL = strSQL & " SELECT TOP 1 mallgubun, makerid, hopesellstat "
	strSQL = strSQL & " FROM db_etcmall.dbo.tbl_jaehumall_hopeSell "
	strSQL = strSQL & " WHERE idx = '"&iidx&"' "
	rsget.Open strSQL,dbget,1
	If Not rsget.Eof then
		fnmallgubun	= rsget("mallgubun")
		fnmakerid	= rsget("makerid")
		fnhopesellstat		= rsget("hopesellstat")
	End If
	rsget.Close

	If fnmallgubun = "all" Then													'제휴사 전체
		strSQL = ""
		strSQL = strSQL & " UPDATE db_user.dbo.tbl_user_c SET " & vbcrlf
		strSQL = strSQL & " isextusing = '"&fnhopesellstat&"' " & vbcrlf
		strSQL = strSQL & " WHERE userid = '"&fnmakerid&"' " & vbcrlf
		dbget.Execute strSQL
		'hopesell의 상태값 변경
		strSQL = ""
		strSQL = strSQL & " UPDATE db_etcmall.dbo.tbl_jaehumall_hopeSell SET "
		strSQL = strSQL & " isComplete = 'Y' "
		strSQL = strSQL & " ,currstat = '3' "
		strSQL = strSQL & " ,adminRegdate = getdate() "
		strSQL = strSQL & " WHERE idx = '"&iidx&"' "
		dbget.Execute strSQL

		If fnhopesellstat = "Y" Then			'전체 Y일 때는 [db_temp].dbo.tbl_jaehyumall_not_in_makerid의 데이터 삭제
			strSQL = ""
			strSQL = strSQL & " DELETE FROM [db_temp].dbo.tbl_jaehyumall_not_in_makerid WHERE makerid = '"&fnmakerid&"' and mallgubun <> 'nvstorefarm' "
			dbget.Execute strSQL

			strSQL = ""
			strSQL = strSQL & " DELETE FROM [db_outmall].dbo.tbl_jaehyumall_not_in_makerid WHERE makerid = '"&fnmakerid&"' and mallgubun <> 'nvstorefarm' "
			dbCTget.Execute strSQL
		ElseIf fnhopesellstat = "N" Then		'전체 N일 때는 [db_temp].dbo.tbl_jaehyumall_not_in_makerid의 데이터 입력
			strSQL = ""
			strSQL = strSQL & " INSERT INTO [db_temp].dbo.tbl_jaehyumall_not_in_makerid"
			strSQL = strSQL & " (makerid,mallgubun,regdate,reguserid)"
			strSQL = strSQL & " SELECT '"&fnmakerid&"', K.mayMallID,getdate(),'"&session("ssBctID")&"'" &VbCRLF
			strSQL = strSQL & " FROM (select c.userid as mayMallID from db_user.dbo.tbl_user_c c JOIN db_partner.dbo.tbl_partner_addInfo f "
			strSQL = strSQL & "       on c.userid=f.partnerid and f.pcomType=1 and f.pmallSellType=1 where c.isusing='Y' and c.userdiv='50') K "
			strSQL = strSQL & " LEFT JOIN [db_temp].dbo.tbl_jaehyumall_not_in_makerid T "
			strSQL = strSQL & " on K.mayMallID =T.mallgubun and T.makerid='"&fnmakerid&"'"
			strSQL = strSQL & " WHERE T.makerid is NULL"
			strSQL = strSQL & " AND K.mayMallID <> 'nvstorefarm' "

			dbget.Execute strSQL

			strSQL = ""
			strSQL = strSQL & " INSERT INTO [db_outmall].dbo.tbl_jaehyumall_not_in_makerid"
			strSQL = strSQL & " (makerid,mallgubun,regdate,reguserid)"
			strSQL = strSQL & " SELECT '"&fnmakerid&"', K.mayMallID,getdate(),'"&session("ssBctID")&"'" &VbCRLF
			strSQL = strSQL & " FROM (select c.userid as mayMallID from db_AppWish.dbo.tbl_user_c c JOIN db_AppWish.dbo.tbl_partner_addInfo f "
			strSQL = strSQL & "       on c.userid=f.partnerid and f.pcomType=1 and f.pmallSellType=1 where c.isusing='Y' and c.userdiv='50') K "
			strSQL = strSQL & " LEFT JOIN [db_outmall].dbo.tbl_jaehyumall_not_in_makerid T "
			strSQL = strSQL & " on K.mayMallID =T.mallgubun and T.makerid='"&fnmakerid&"'"
			strSQL = strSQL & " WHERE T.makerid is NULL"
			strSQL = strSQL & " AND K.mayMallID <> 'nvstorefarm' "
			dbCTget.Execute strSQL
		End If

		'로그에 쌓기
		strSQL = ""
		strSQL = strSQL & " INSERT INTO db_etcmall.dbo.tbl_jaehumall_hopeSell_Log (mallgubun, makerid, hopeStr, useYN, reguserid, regdate) " & vbcrlf
		strSQL = strSQL & " SELECT TOP 1 mallgubun, makerid, '[관리자] 승인 완료', hopesellstat, '"&session("ssBctID")&"', getdate() "
		strSQL = strSQL & " FROM db_etcmall.dbo.tbl_jaehumall_hopeSell "
		strSQL = strSQL & " WHERE idx= '"&iidx&"' " & vbcrlf
		dbget.Execute strSQL
	ElseIf (fnmallgubun = "nvstorefarm") OR (fnOverseasMall(fnmallgubun) = "Y") Then					'user_c의 isextusing을 건드리지 않고 판매하는 경우 (스토어팜)
	'fnOverseasMall(fnmallgubun) 조건 추가..해외몰 이면 isextusing을 건드리지 않음
		If fnhopesellstat = "N" Then
			strSQL = "IF NOT Exists(select * from [db_temp].dbo.tbl_jaehyumall_not_in_makerid where mallgubun='"&fnmallgubun&"' and makerid='"&fnmakerid&"')"
			strSQL = strSQL&" BEGIN"
			strSQL = strSQL&" Insert into [db_temp].dbo.tbl_jaehyumall_not_in_makerid "
			strSQL = strSQL&" (makerid,mallgubun,regdate,reguserid)"
			strSQL = strSQL&" values('"&fnmakerid&"','"&fnmallgubun&"',getdate(),'"&session("ssBctID")&"')"
			strSQL = strSQL&" END "
			dbget.Execute strSQL

			strSQL = "IF NOT Exists(select * from [db_outmall].dbo.tbl_jaehyumall_not_in_makerid where mallgubun='"&fnmallgubun&"' and makerid='"&fnmakerid&"')"
			strSQL = strSQL&" BEGIN"
			strSQL = strSQL&" Insert into [db_outmall].dbo.tbl_jaehyumall_not_in_makerid "
			strSQL = strSQL&" (makerid,mallgubun,regdate,reguserid)"
			strSQL = strSQL&" values('"&fnmakerid&"','"&fnmallgubun&"',getdate(),'"&session("ssBctID")&"')"
			strSQL = strSQL&" END "
			dbCTget.Execute strSQL
		Else                              ''등록가능
			strSQL = "IF Exists(select * from [db_temp].dbo.tbl_jaehyumall_not_in_makerid where mallgubun='"&fnmallgubun&"' and makerid='"&fnmakerid&"')"
			strSQL = strSQL&" BEGIN"
			strSQL = strSQL&" delete from [db_temp].dbo.tbl_jaehyumall_not_in_makerid where mallgubun='"&fnmallgubun&"' and makerid='"&fnmakerid&"'"
			strSQL = strSQL&" END "
			dbget.Execute strSQL

			strSQL = "IF Exists(select * from [db_outmall].dbo.tbl_jaehyumall_not_in_makerid where mallgubun='"&fnmallgubun&"' and makerid='"&fnmakerid&"')"
			strSQL = strSQL&" BEGIN"
			strSQL = strSQL&" delete from [db_outmall].dbo.tbl_jaehyumall_not_in_makerid where mallgubun='"&fnmallgubun&"' and makerid='"&fnmakerid&"'"
			strSQL = strSQL&" END "
			dbCTget.Execute strSQL
		End If
		strSQL = ""
		strSQL = strSQL & " UPDATE db_etcmall.dbo.tbl_jaehumall_hopeSell SET "
		strSQL = strSQL & " isComplete = 'Y' "
		strSQL = strSQL & " ,currstat = '3' "
		strSQL = strSQL & " ,adminRegdate = getdate() "
		strSQL = strSQL & " WHERE idx = '"&iidx&"' "
		dbget.Execute strSQL

		'로그에 쌓기
		strSQL = ""
		strSQL = strSQL & " INSERT INTO db_etcmall.dbo.tbl_jaehumall_hopeSell_Log (mallgubun, makerid, hopeStr, useYN, reguserid, regdate) " & vbcrlf
		strSQL = strSQL & " SELECT TOP 1 mallgubun, makerid, '[관리자] 승인 완료', hopesellstat, '"&session("ssBctID")&"', getdate() "
		strSQL = strSQL & " FROM db_etcmall.dbo.tbl_jaehumall_hopeSell "
		strSQL = strSQL & " WHERE idx= '"&iidx&"' " & vbcrlf
		dbget.Execute strSQL
	ElseIf fnmallgubun = "naverep" OR fnmallgubun = "daumep" OR fnmallgubun = "shodocep" OR fnmallgubun = "wemakepriceep" OR fnmallgubun = "ggshop" Then					'다음 OR 네이버	OR 쇼닥
		strSQL = ""
		strSQL = strSQL & " IF NOT Exists(select * from db_temp.dbo.tbl_EpShop_not_in_makerid where mallgubun='"&fnmallgubun&"' and makerid='"&fnmakerid&"')"
		strSQL = strSQL & " BEGIN"
		strSQL = strSQL & " 	INSERT INTO db_temp.dbo.tbl_EpShop_not_in_makerid (makerid, mallgubun, isusing, regdate, regid) VALUES "
		strSQL = strSQL & " 	('"&fnmakerid&"', '"&fnmallgubun&"', '"&fnhopesellstat&"' ,getdate(), '"&session("ssBctID")&"') "
        strSQL = strSQL & " END Else "
		strSQL = strSQL & " BEGIN"
		strSQL = strSQL & "		UPDATE db_temp.dbo.tbl_EpShop_not_in_makerid SET "
		strSQL = strSQL & " 	isusing = '"&fnhopesellstat&"'"
		strSQL = strSQL & " 	,lastupdate = getdate()"
		strSQL = strSQL & " 	,updateid = '"&session("ssBctID")&"'"
		strSQL = strSQL & " 	WHERE makerid = '"&fnmakerid&"' "
		strSQL = strSQL & " 	AND mallgubun = '"&fnmallgubun&"' "
        strSQL = strSQL & " END "
        dbget.Execute strSQL

		strSQL = ""
		strSQL = strSQL & " IF NOT Exists(select * from db_outmall.dbo.tbl_EpShop_not_in_makerid where mallgubun='"&fnmallgubun&"' and makerid='"&fnmakerid&"')"
		strSQL = strSQL & " BEGIN"
		strSQL = strSQL & " 	INSERT INTO db_outmall.dbo.tbl_EpShop_not_in_makerid (makerid, mallgubun, isusing, regdate, regid) VALUES "
		strSQL = strSQL & " 	('"&fnmakerid&"', '"&fnmallgubun&"', '"&fnhopesellstat&"' ,getdate(), '"&fnmakerid&"') "
        strSQL = strSQL & " END Else "
		strSQL = strSQL & " BEGIN"
		strSQL = strSQL & " 	UPDATE db_outmall.dbo.tbl_EpShop_not_in_makerid SET "
		strSQL = strSQL & " 	isusing = '"&fnhopesellstat&"'"
		strSQL = strSQL & " 	,lastupdate = getdate()"
		strSQL = strSQL & " 	,updateid = '"&session("ssBctID")&"'"
		strSQL = strSQL & " 	WHERE makerid = '"&fnmakerid&"' "
		strSQL = strSQL & " 	AND mallgubun = '"&fnmallgubun&"' "
        strSQL = strSQL & " END "
        dbCTget.Execute strSQL

		strSQL = ""
		strSQL = strSQL & " UPDATE db_etcmall.dbo.tbl_jaehumall_hopeSell SET "
		strSQL = strSQL & " isComplete = 'Y' "
		strSQL = strSQL & " ,currstat = '3' "
		strSQL = strSQL & " ,adminRegdate = getdate() "
		strSQL = strSQL & " WHERE idx = '"&iidx&"' "
		dbget.Execute strSQL

		'로그에 쌓기
		strSQL = ""
		strSQL = strSQL & " INSERT INTO db_etcmall.dbo.tbl_jaehumall_hopeSell_Log (mallgubun, makerid, hopeStr, useYN, reguserid, regdate) " & vbcrlf
		strSQL = strSQL & " SELECT TOP 1 mallgubun, makerid, '[관리자] 승인 완료', hopesellstat, '"&session("ssBctID")&"', getdate() "
		strSQL = strSQL & " FROM db_etcmall.dbo.tbl_jaehumall_hopeSell "
		strSQL = strSQL & " WHERE idx= '"&iidx&"' " & vbcrlf
		dbget.Execute strSQL
	Else																		'그 외
		If fnhopesellstat = "N" Then
			strSQL = "IF NOT Exists(select * from [db_temp].dbo.tbl_jaehyumall_not_in_makerid where mallgubun='"&fnmallgubun&"' and makerid='"&fnmakerid&"')"
			strSQL = strSQL&" BEGIN"
			strSQL = strSQL&" Insert into [db_temp].dbo.tbl_jaehyumall_not_in_makerid "
			strSQL = strSQL&" (makerid,mallgubun,regdate,reguserid)"
			strSQL = strSQL&" values('"&fnmakerid&"','"&fnmallgubun&"',getdate(),'"&session("ssBctID")&"')"
			strSQL = strSQL&" END "
			dbget.Execute strSQL

			strSQL = "IF NOT Exists(select * from [db_outmall].dbo.tbl_jaehyumall_not_in_makerid where mallgubun='"&fnmallgubun&"' and makerid='"&fnmakerid&"')"
			strSQL = strSQL&" BEGIN"
			strSQL = strSQL&" Insert into [db_outmall].dbo.tbl_jaehyumall_not_in_makerid "
			strSQL = strSQL&" (makerid,mallgubun,regdate,reguserid)"
			strSQL = strSQL&" values('"&fnmakerid&"','"&fnmallgubun&"',getdate(),'"&session("ssBctID")&"')"
			strSQL = strSQL&" END "
			dbCTget.Execute strSQL
		Else                              ''등록가능

			strSQL = ""
			strSQL = strSQL & " SELECT TOP 1 isextusing "
			strSQL = strSQL & " FROM db_user.dbo.tbl_user_c "
			strSQL = strSQL & " WHERE userid = '"&fnmakerid&"' " & vbcrlf
			rsget.Open strSQL,dbget,1
			If Not rsget.Eof then
				currExtusing = rsget("isextusing")
			End If
			rsget.Close

			if (currExtusing = "N") then
				'// 제휴사 전체 판매안함 상태에서 특정 브랜드판매함 전환인 경우
				strSQL = ""
				strSQL = strSQL & " UPDATE db_user.dbo.tbl_user_c SET " & vbcrlf
				strSQL = strSQL & " isextusing = 'Y' " & vbcrlf
				strSQL = strSQL & " WHERE userid = '"&fnmakerid&"' " & vbcrlf
				dbget.Execute strSQL

				strSQL = ""
				strSQL = strSQL & " INSERT INTO [db_temp].dbo.tbl_jaehyumall_not_in_makerid"
				strSQL = strSQL & " (makerid,mallgubun,regdate,reguserid)"
				strSQL = strSQL & " SELECT '"&fnmakerid&"', K.mayMallID,getdate(),'"&session("ssBctID")&"'" &VbCRLF
				strSQL = strSQL & " FROM (select c.userid as mayMallID from db_user.dbo.tbl_user_c c JOIN db_partner.dbo.tbl_partner_addInfo f "
				strSQL = strSQL & "       on c.userid=f.partnerid and f.pcomType=1 and f.pmallSellType=1 where c.isusing='Y' and c.userdiv='50') K "
				strSQL = strSQL & " LEFT JOIN [db_temp].dbo.tbl_jaehyumall_not_in_makerid T "
				strSQL = strSQL & " on K.mayMallID =T.mallgubun and T.makerid='"&fnmakerid&"'"
				strSQL = strSQL & " WHERE T.makerid is NULL"
				dbget.Execute strSQL

				strSQL = ""
				strSQL = strSQL & " INSERT INTO [db_outmall].dbo.tbl_jaehyumall_not_in_makerid"
				strSQL = strSQL & " (makerid,mallgubun,regdate,reguserid)"
				strSQL = strSQL & " SELECT '"&fnmakerid&"', K.mayMallID,getdate(),'"&session("ssBctID")&"'" &VbCRLF
				strSQL = strSQL & " FROM (select c.userid as mayMallID from db_AppWish.dbo.tbl_user_c c JOIN db_AppWish.dbo.tbl_partner_addInfo f "
				strSQL = strSQL & "       on c.userid=f.partnerid and f.pcomType=1 and f.pmallSellType=1 where c.isusing='Y' and c.userdiv='50') K "
				strSQL = strSQL & " LEFT JOIN [db_outmall].dbo.tbl_jaehyumall_not_in_makerid T "
				strSQL = strSQL & " on K.mayMallID =T.mallgubun and T.makerid='"&fnmakerid&"'"
				strSQL = strSQL & " WHERE T.makerid is NULL"
				dbCTget.Execute strSQL
			end if

			strSQL = "IF Exists(select * from [db_temp].dbo.tbl_jaehyumall_not_in_makerid where mallgubun='"&fnmallgubun&"' and makerid='"&fnmakerid&"')"
			strSQL = strSQL&" BEGIN"
			strSQL = strSQL&" delete from [db_temp].dbo.tbl_jaehyumall_not_in_makerid where mallgubun='"&fnmallgubun&"' and makerid='"&fnmakerid&"'"
			strSQL = strSQL&" END "
			dbget.Execute strSQL

			strSQL = "IF Exists(select * from [db_outmall].dbo.tbl_jaehyumall_not_in_makerid where mallgubun='"&fnmallgubun&"' and makerid='"&fnmakerid&"')"
			strSQL = strSQL&" BEGIN"
			strSQL = strSQL&" delete from [db_outmall].dbo.tbl_jaehyumall_not_in_makerid where mallgubun='"&fnmallgubun&"' and makerid='"&fnmakerid&"'"
			strSQL = strSQL&" END "
			dbCTget.Execute strSQL
		End If

		strSQL = ""
		strSQL = strSQL & " UPDATE db_etcmall.dbo.tbl_jaehumall_hopeSell SET "
		strSQL = strSQL & " isComplete = 'Y' "
		strSQL = strSQL & " ,currstat = '3' "
		strSQL = strSQL & " ,adminRegdate = getdate() "
		strSQL = strSQL & " WHERE idx = '"&iidx&"' "
		dbget.Execute strSQL

		'로그에 쌓기
		strSQL = ""
		strSQL = strSQL & " INSERT INTO db_etcmall.dbo.tbl_jaehumall_hopeSell_Log (mallgubun, makerid, hopeStr, useYN, reguserid, regdate) " & vbcrlf
		strSQL = strSQL & " SELECT TOP 1 mallgubun, makerid, '[관리자] 승인 완료', hopesellstat, '"&session("ssBctID")&"', getdate() "
		strSQL = strSQL & " FROM db_etcmall.dbo.tbl_jaehumall_hopeSell "
		strSQL = strSQL & " WHERE idx= '"&iidx&"' " & vbcrlf
		dbget.Execute strSQL
	End If
End Function
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbCTclose.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->
