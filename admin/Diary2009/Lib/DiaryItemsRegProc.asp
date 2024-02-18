<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/admin/diary2009/lib/include_event_code.asp"-->
<%
'// 다이어리 멀티 저장 처리 페이지 2018-08-17 이종화
dim i
dim CateCode , mode
dim idx , tempidx
dim itemcount
dim strSQL
dim idxStrSQL

mode = request.Form("mode")
CateCode = request.Form("cate")
itemcount = request.Form("chkitem").count

if CateCode = "" then CateCode = 0
	
IF mode="I" Then
        dbget.beginTrans
    For i = 1 To itemcount	'파일갯수 만큼 업로드
	    strSQL =" INSERT INTO db_diary2010.[dbo].tbl_DiaryMaster " & vbcrlf
        strSQL = strSQL & " (Cate,Itemid,isusing,commentyn,event_code,eventgroup_code,comment_img ,weight, mdpick, limited, storytext , mdpicksort, event_start, event_end) " & vbcrlf
        strSQL = strSQL & " VALUES("  & vbcrlf
        strSQL = strSQL & "'" & CateCode & "' "  & vbcrlf
        strSQL = strSQL & ",'" & request.Form("chkitem")(i) & "' "  & vbcrlf
        strSQL = strSQL & ",'Y' "  & vbcrlf
        strSQL = strSQL & ",'' "  & vbcrlf
        strSQL = strSQL & ",'0' "  & vbcrlf
        strSQL = strSQL & ",'0' "  & vbcrlf
        strSQL = strSQL & ",'' "  & vbcrlf
        strSQL = strSQL & ",'0' "  & vbcrlf
        strSQL = strSQL & ",'x' "  & vbcrlf
        strSQL = strSQL & ",'x' "  & vbcrlf
        strSQL = strSQL & ",''"  & vbcrlf
        strSQL = strSQL & ",'0' "  & vbcrlf
	    strSQL = strSQL & ",null "  & vbcrlf
        strSQL = strSQL & ",null "  & vbcrlf
        strSQL = strSQL & " )"

	    'response.write strSQL&"<br>"
	    dbget.execute(strSQL)

        idxStrSQL = "SELECT SCOPE_IDENTITY()"
        rsget.open idxStrSQL,dbget,2
        IF not rsget.Eof Then
            tempidx = rsget(0)
        End IF
        rsget.close

	    idx = tempidx

        '2019 다이어리 내지 구성
        strSQL = " INSERT INTO [db_diary2010].[dbo].tbl_diary_Info (idx ,info_gubun,info_name,info_pageCnt) " & vbcrlf
        strSQL = strSQL & " VALUES " & vbcrlf
        'strSQL = strSQL & "('" & idx & "','22','1개월','0') ," & vbcrlf
        strSQL = strSQL & "('" & idx & "','23','분기별','0') ," & vbcrlf
        strSQL = strSQL & "('" & idx & "','24','6개월','0') ," & vbcrlf
        strSQL = strSQL & "('" & idx & "','25','1년','0') ," & vbcrlf
        strSQL = strSQL & "('" & idx & "','26','1년 이상','0') ," & vbcrlf
        strSQL = strSQL & "('" & idx & "','27','연간스케줄','0') ," & vbcrlf
        'strSQL = strSQL & "('" & idx & "','28','월간스케줄','0') ," & vbcrlf
        'strSQL = strSQL & "('" & idx & "','29','주간스케줄','0') ," & vbcrlf
        'strSQL = strSQL & "('" & idx & "','30','일스케줄','0') ," & vbcrlf
        'strSQL = strSQL & "('" & idx & "','31','캐시북','0') ," & vbcrlf
        strSQL = strSQL & "('" & idx & "','32','포켓','0') ," & vbcrlf
        strSQL = strSQL & "('" & idx & "','33','밴드','0') ," & vbcrlf
        strSQL = strSQL & "('" & idx & "','34','펜홀더','0') ," & vbcrlf
        strSQL = strSQL & "('" & idx & "','35','만년형','0') ," & vbcrlf
        strSQL = strSQL & "('" & idx & "','36','2019 날짜형','0') ," & vbcrlf
        strSQL = strSQL & "('" & idx & "','37','먼슬리','0') ," & vbcrlf
        strSQL = strSQL & "('" & idx & "','38','위클리','0') ," & vbcrlf
        strSQL = strSQL & "('" & idx & "','39','데일리','0') ," & vbcrlf
        strSQL = strSQL & "('" & idx & "','40','다이어리','0') ," & vbcrlf
        strSQL = strSQL & "('" & idx & "','41','스터디','0') ," & vbcrlf
        strSQL = strSQL & "('" & idx & "','42','가계부','0') ," & vbcrlf
        strSQL = strSQL & "('" & idx & "','43','자기계발','0') " & vbcrlf	
        dbget.execute(strSQL)
    Next 
	'오류검사 및 반영
	If Err.Number = 0 Then
		dbget.CommitTrans				'커밋(정상)
        Alert_return "저장 되었습니다."
        response.write "<script type='text/javascript'>parent.window.close();</script>"
	Else
		dbget.RollBackTrans				'롤백(에러발생시)
		Alert_return "처리중 에러가 발생했습니다."
	End If
End IF
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->