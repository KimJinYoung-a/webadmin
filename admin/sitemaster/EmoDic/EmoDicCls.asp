<%

Class EmodicCls
	
	private Sub Class_initialize()
	
	End Sub
	
	Private Sub Class_terminate()
	
	End Sub
	
	dim FRectEmoNumber
	dim FRectEmoType
	dim FRectEmoTitle
	
	dim FList
	dim FResultCount 
	
	
	Sub getEmoWordsList()
		dim strSQL,i
		
		strSQL =" SELECT EmoNO,EmoType,EmoDesc,EmoTitle,EmoImage,EmoUsing,EmoSortNo " &_
				" FROM [db_sitemaster].[dbo].[tbl_emowords] " &_
				" WHERE EmoNo="& FRectEmoNumber & " and EmoType=" & FRectEmoType
				
				IF FRectEmoTitle<>"" Then
					strSQL= strSQL& " and EmoTitle='"& FRectEmoTitle &"'"
				End if
				strSQL= strSQL &" ORDER BY EmoSortNo"
		
		'response.write strSQL		
		rsget.open strSQL,dbget,1
		
		FResultCount = rsget.RecordCount
		
		'response.write FResultCount
		redim FList(FResultCount)
		
		IF not rsget.eof then
			Do until rsget.eof 
			
			Set FList(i) = new EmoDicWordsCls
			
				FList(i).EmoNO = rsget("EmoNO")
				FList(i).EmoType	= rsget("EmoType")
				FList(i).EmoDesc	= rsget("EmoDesc")
				FList(i).EmoTitle	= rsget("EmoTitle")
				FList(i).EmoImage 	= rsget("EmoImage")
				FList(i).EmoUsing 	= rsget("EmoUsing")
				FList(i).EmoSortNO	= rsget("EmoSortNo")
				i= i+1
				rsget.moveNext
			Loop
				
		End IF	
		rsget.close
	
	End Sub
	
	'//단어별 코멘트 리스트 
	Sub getEmoCommentList
		dim strSQL,i
		strSQL =" SELECT idx,ecUserid,ecComment " &_
				" FROM [db_sitemaster].[dbo].[tbl_emowords_comment] " &_
				" WHERE 1=1 and ecUsing='Y'" &_
				" and EmoNO = "& FRectEmoNumber &_
				" and EmoType ="& FRectEmoType &_
				" and EmoTitle='"& FRectEmoTitle &"'" &_
				" ORDER BY idx desc "
				
		rsget.open strSQL,dbget,1
		
		'response.write strSQL
		
		FResultCount = rsget.RecordCount
		'response.write FResultCount
		
		redim FList(FResultCount)
		
		IF not rsget.eof then
			Do until rsget.eof 
			
			Set FList(i) = new EmoDicCommentsCls
			
				FList(i).EmoIdx = rsget("idx")
				FList(i).Userid = rsget("ecUserid")
				FList(i).Comment = db2html(rsget("ecComment"))
				'FList(i).EmoType	= rsget("EmoType")
				'FList(i).EmoDesc	= db2html(rsget("EmoDesc"))
				'FList(i).EmoTitle	= db2html(rsget("EmoTitle"))
				'FList(i).EmoImage 	= db2html(rsget("EmoImage"))
				'FList(i).EmoUsing 	= rsget("EmoUsing")
				
				i= i+1
				rsget.moveNext
			Loop
			
		End IF	
		rsget.close
		
	End Sub
End Class

Class EmoDicWordsCls

	private Sub Class_initialize()
	
	End Sub
	
	Private Sub Class_terminate()
	
	End Sub
	
	dim EmoNO
	dim EmoType
	dim EmoDesc
	dim EmoTitle
	dim EmoImage
	dim EmoUsing
	dim EmoSortNO
	
	Function getImgUrl()
		IF application("Svr_Info")="Dev" Then
			getImgUrl = "<img src=http://testimgstatic.10x10.co.kr/contents/Emodic/"&EmoImage&" border=""0"">"
		ELSE 
			getImgUrl = "<img src=""http://imgstatic.10x10.co.kr/contents/Emodic/"&EmoImage&""" border=""0"">"
		End if
	End Function
	
End Class 

Class EmoDicCommentsCls

	private Sub Class_initialize()
	
	End Sub
	
	Private Sub Class_terminate()
	
	End Sub
	dim EmoIdx
	dim EmoNo
	dim EmoType
	dim Userid
	dim Comment
	dim RegDate
	
	
End Class 


%>