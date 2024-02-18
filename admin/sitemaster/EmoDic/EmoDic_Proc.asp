<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/sitemaster/EmoDic/EmoDicCls.asp" -->


<%

dim eNumber,eType,mode,iLp,strSQL
eNumber = request("eno")
eType = request("etp")
mode=request("mode")

dim isUsing 
isUsing = request("ius")

dim arrWords
arrWords = request("awrd")

dim arrSortNo
arrSortNo = request("srtno")

IF (eNumber="") or (eType="") or (mode="") Then
	response.write "잘못된 접근입니다"
	dbget.close()	:	response.End
End if



On Error Resume Next 

IF mode="batch" Then ''최초 일괄등록시 
	
	dim strCkSQL
	
	
	arrWords = split(arrWords,",")
	
	For iLp = 0 To Ubound(arrWords)
		
		strCkSQL =" SELECT Count(*) as cnt FROM [db_sitemaster].[dbo].[tbl_emowords] WHERE EmoNo="& eNumber &" and EmoType="& eType &" and EmoTitle='" & trim(arrWords(iLp)) & "'"
		
		rsget.open strCkSQL,dbget,1
		
		IF 	rsget("cnt")=0 THEN
			
		strSQL = strSQL &_
			
			" INSERT INTO [db_sitemaster].[dbo].tbl_emowords (EmoNo,EmoType,EmoTitle,EmoSortNo) " &_
			" Values ("& eNumber &","& eType &",'"& trim(arrWords(iLp)) &"'," & iLp &")"
			'" Values ("& eNumber &","& eType &","& EmoTitle &","& EmoDesc &"," & EmoImage &")"
		End IF
		rsget.Close
		
	Next 
	
ELSEIF mode="allUsing" Then
	strSQL =" UPDATE [db_sitemaster].[dbo].tbl_emowords" &_
			" SET EmoUsing='"& isUsing &"'" &_
			" WHERE EmoNo="&eNumber&" and EmoType="&eType
			

ELSEIF mode="arrEdit" Then
	dim arrisUsing,tmp
	
	arrWords = split(arrWords,",")
	arrSortNo = split(arrSortNo,",")
	arrIsUsing = split(left(isUsing,len(isUsing)-1),",")
	
	
	
	IF (isArray(arrWords) and isArray(arrSortNo) and isArray(arrIsUsing)) THEN
		
		FOR iLp=0 To Ubound(arrWords)
		
		tmp =" UPDATE [db_sitemaster].[dbo].tbl_emowords "&_ 
			" SET EmoUsing='"& arrisUsing(iLp) &"'" &_
			" ,EmoSortNo ="& arrSortNo(iLp) &"" &_
			" Where EmoNo="& eNumber &_
			" and EmoType="& eType &_
			" and EmoTitle='"& trim(arrWords(iLp)) &"' "
		
		strSQL = strSQL & tmp		
		'response.write strSQL
				
		
		NEXT
	ELSE
		
	End IF

	
ELSEIF mode="del" Then

END IF

'response.write strSQL
'dbget.close()	:	response.End
dbget.BeginTrans

dbget.Execute(strSQL)


IF Err.Number= 0 Then
	dbget.CommitTrans
	
	IF mode="allUsing" Then
		response.write "<script>alert('저장되었습니다.'); parent.location='EmodicManage.asp?eno="&eNumber&"&etp="&eType&"';</script>"
	Else
		response.write "<script>alert('저장되었습니다.'); parent.location='EmodicManage.asp?eno="&eNumber&"&etp="&eType&"'; self.close();</script>"
	End if
	
	dbget.close()	:	response.End
Else	
	dbget.RollbackTrans
	response.write "<script>alert('에러발생\n다시 입렵해주소 ㅡ,.ㅡㅋ.'); history.go(-1);</script>"
	dbget.close()	:	response.End	
End IF


%>


<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->