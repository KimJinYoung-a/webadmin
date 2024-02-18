<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/etc/xSiteTempOrderCls.asp"-->

<%
	Dim vQuery, vOutMallOrderSeq, vMatchItemID, vMatchItemOption, vIsOK, mode
	
	vOutMallOrderSeq	= requestCheckvar(request("outMallorderSeq"),32)
	vMatchItemID		= requestCheckvar(request("Matchitemid"),32)
	vMatchItemOption	= requestCheckvar(request("matchitemoption"),32)
	mode                = requestCheckvar(request("mode"),32)
	if (mode="optmatch") then
    	vQuery = "" & _
    			"IF EXISTS(SELECT matchItemID FROM [db_temp].[dbo].[tbl_xSite_TMPOrder] WHERE OutMallOrderSeq = '" & vOutMallOrderSeq & "' AND MatchItemID = '" & vMatchItemID & "') " & _
    			"	BEGIN " & _
    			"		SELECT 'o' " & _
    			"		UPDATE [db_temp].[dbo].[tbl_xSite_TMPOrder] SET " & _
    			"			matchitemoption = '" & vMatchItemOption & "' " & _
    			"		WHERE OutMallOrderSeq = '" & vOutMallOrderSeq & "' AND MatchItemID = '" & vMatchItemID & "' " & _
    			"	 " & _
    			"	END " & _
    			"ELSE " & _
    			"	BEGIN " & _
    			"		SELECT 'x' " & _
    			"	END "
    	rsget.open vQuery, dbget,1

    	vIsOK = "x"
    	If Not rsget.Eof Then
    		IF rsget(0) = "o" Then
    			vIsOK = "o"
    		End If
    	End IF
		vQuery = ""
		vQuery = vQuery & " INSERT INTO db_temp.[dbo].[tbl_xSite_TmpOrder_ModifyOption] (OutmallOrderSeq, MatchItemID, matchitemoption, userId, regdate) VALUES "
		vQuery = vQuery & " ('"& vOutMallOrderSeq &"', '" & vMatchItemID & "', '"& vMatchItemOption &"', '"& session("ssBctID") &"', GETDATE()) "
		dbget.execute(vQuery)
    elseif (mode="optnone") then
        vQuery = "" & _
    			"IF EXISTS(SELECT matchItemID FROM [db_temp].[dbo].[tbl_xSite_TMPOrder] WHERE OutMallOrderSeq = '" & vOutMallOrderSeq & "' AND MatchItemID = '" & vMatchItemID & "') " & _
    			"	BEGIN " & _
    			"		SELECT 'o' " & _
    			"		UPDATE [db_temp].[dbo].[tbl_xSite_TMPOrder] SET " & _
    			"			matchitemoption = '0000' " & _
    			"		WHERE OutMallOrderSeq = '" & vOutMallOrderSeq & "' AND MatchItemID = '" & vMatchItemID & "' " & _
    			"	 " & _
    			"	END " & _
    			"ELSE " & _
    			"	BEGIN " & _
    			"		SELECT 'x' " & _
    			"	END "
    	rsget.open vQuery, dbget,1

    	vIsOK = "x"
    	If Not rsget.Eof Then
    		IF rsget(0) = "o" Then
    			vIsOK = "o"
    		End If
    	End IF
    end if
%>

<script language="javascript">
<% If vIsOK = "o" Then %>
	opener.document.location.reload();
	window.close();
<% Else %>
	alert("해당 데이터가 존재하지 않거나 옳바른 값이 아닙니다.\n아래 내용을 메모해두었다가 시스템팀에 문의해주세요.\n\noutMallorderSeq = <%=vOutMallOrderSeq%>, Matchitemid = <%=vMatchItemID%>, matchitemoption = <%=vMatchItemOption%>");
	history.back();
<% End If %>
</script>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->