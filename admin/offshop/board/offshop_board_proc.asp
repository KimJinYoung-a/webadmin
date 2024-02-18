<%@ codepage="65001" language="VBScript" %>
<% option explicit %>
<% response.Charset="UTF-8" %>
<%
session.codePage = 65001
%>
<%
	Response.AddHeader "Cache-Control","no-cache"
	Response.AddHeader "Expires","0"
	Response.AddHeader "Pragma","no-cache"
%>
<%
'###########################################################
' Description :  오프라인 통합 게시판
' History : 2010.06.18 한용민 생성
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/common/checkPoslogin.asp"-->
<!-- #include virtual="/common/incSessionAdminorShop.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/common/lib/commonbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/offshop/board/board_cls_utf8.asp"-->

<%
Dim strSql, vWorkerTemp, vWorkerViewTemp, vFileTemp, i , g_MenuPos , mode ,iDoc_Idx
Dim sDoc_Id, sDoc_Name, sDoc_Status,sDoc_Type, sDoc_Import, sDoc_Diffi, sDoc_Subj
dim sDoc_Content, sDoc_Worker, sDoc_File, sDoc_WorkerView, sDoc_UseYN
dim dispshopall ,dispshopdivon ,dispshopdiv ,dispshopidon ,shopid ,doc_kind
	iDoc_Idx		= requestCheckVar(Request("didx"),10)
	sDoc_Id			= session("ssBctId")
	sDoc_UseYN		= requestCheckVar(Request("doc_useyn"),1)
	sDoc_Status		= requestCheckVar(Request("K000"),24)	
	sDoc_Type		= requestCheckVar(Request("G000"),24)
	sDoc_Import		= requestCheckVar(Request("L000"),24)
	sDoc_Diffi		= requestCheckVar(Request("doc_difficult"),2)
	sDoc_Worker		= requestCheckVar(Request("doc_worker"),1000)
	sDoc_Subj		= requestCheckVar(Request("doc_subject"),150)
	sDoc_Content	= replace(Request("doc_content"),"'","")
	sDoc_File		= requestCheckVar(Request("doc_file"),150)
	mode		=  requestCheckVar(Request("mode"),32)
	g_MenuPos =  requestCheckVar(request("menupos"),10)
	dispshopdivon =  requestCheckVar(request("dispshopdivon"),2)
	dispshopdiv =  requestCheckVar(request("A000"),2)
	dispshopall =  requestCheckVar(request("dispshopall"),1)
	dispshopidon =  requestCheckVar(request("dispshopidon"),2)
	shopid =  request("shopid")
	doc_kind =  requestCheckVar(request("doc_kind"),24)

	'response.write sDoc_Status
	'response.end

if mode = "edit" then
	if sDoc_Subj <> "" and not(isnull(sDoc_Subj)) then
		sDoc_Subj = ReplaceBracket(sDoc_Subj)
	end If
	if sDoc_Content <> "" and not(isnull(sDoc_Content)) then
		sDoc_Content = ReplaceBracket(sDoc_Content)
	end If
	if sDoc_Content <> "" then
		if checkNotValidHTML(sDoc_Content) then
		response.write "<script type='text/javascript'>"
		response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');"
		response.write "</script>"
		dbget.close()	:	response.End
		end if
	end if

	'//신규저장	
	If iDoc_Idx = "" Then
		strSql = " INSERT INTO db_shop.dbo.tbl_offshop_board_document" & VbCRLF
		strSql = strSql & " (id, doc_type, doc_important, doc_difficult, doc_subject, doc_content, doc_status" & VbCRLF
		strSql = strSql & " ,dispshopall , dispshopdiv,doc_kind) VALUES (" & VbCRLF
		strSql = strSql & " '" & sDoc_Id & "','" & sDoc_Type & "', '" & sDoc_Import & "', '" & sDoc_Diffi & "', " & VbCRLF
		strSql = strSql & " '" & html2db(sDoc_Subj) & "', '" & html2db(replace(sDoc_Content,vbcrlf,"")) & "', '" & sDoc_Status & "'" & VbCRLF
				 
		if dispshopall = "ON" then
			strSql = strSql & " ,'Y'" & VbCRLF
		else
			strSql = strSql & " ,NULL" & VbCRLF
		end if
		
		if dispshopdivon = "ON" then
			strSql = strSql & " ,'"&dispshopdiv&"'" & VbCRLF
		else
			strSql = strSql & " ,NULL" & VbCRLF
		end if
		
		strSql = strSql & " ,'"&doc_kind&"'" & VbCRLF		
		strSql = strSql & " )"
		
		'response.write strSql &"<br>"
		dbget.execute strSql
    	
		strSql = ""
		strSql = " SELECT top 1 doc_idx from db_shop.dbo.tbl_offshop_board_document"
		strSql = strSql & " where doc_useyn = 'Y'"
		strSql = strSql & " order by doc_idx desc"
		
		'response.write strSql &"<br>"
		rsget.Open strSql,dbget
		
		IF Not rsget.EOF THEN
			'response.write rsget(0) &"<br>"
			iDoc_Idx = rsget(0)
		ELSE	
			Call sbAlertMsg ("데이터 처리에 문제가 발생하였습니다.[1]", "back", "")
			session.codePage = 949
			rsget.close : dbget.close : response.end
		END IF
		rsget.close

		'/매장등록
		if dispshopidon = "ON" then
			if shopid <> "" then
				shopid = split(shopid,",")
				
				if isarray(shopid) then
					for i = 0 to ubound(shopid)
					
			        strSql = " insert into db_shop.dbo.tbl_offshop_board_shop"&VbCRLF
			        strSql = strSql & " (doc_idx,shopid) values"&VbCRLF
			        strSql = strSql & " ("& iDoc_Idx &",'"& requestCheckVar(trim(shopid(i)),32) &"')"
			        
					'response.write strSql &"<br>"
					dbget.execute strSql
					
					next
				end if
		    end if
    	end if
	
	'//수정	
	Else
		strSql = " UPDATE db_shop.dbo.tbl_offshop_board_document SET" & VbCRLF
		strSql = strSql & " doc_type = '" & sDoc_Type & "'," & VbCRLF
		strSql = strSql & " doc_important = '" & sDoc_Import & "'," & VbCRLF
		strSql = strSql & " doc_difficult = '" & sDoc_Diffi & "'," & VbCRLF
		strSql = strSql & " doc_subject = '" & sDoc_Subj & "'," & VbCRLF
		strSql = strSql & " doc_content = '" & html2db(replace(sDoc_Content,vbcrlf,"")) & "'," & VbCRLF
		strSql = strSql & " doc_status = '" & sDoc_Status & "'," & VbCRLF
		strSql = strSql & " doc_useyn = '" & sDoc_UseYN & "'," & VbCRLF
				 
		if dispshopall = "ON" then
			strSql = strSql & " dispshopall = 'Y'," & VbCRLF
		else
			strSql = strSql & " dispshopall = NULL," & VbCRLF
		end if
		
		if dispshopdivon = "ON" then
			strSql = strSql & " dispshopdiv = '" & dispshopdiv & "'," & VbCRLF
		else
			strSql = strSql & " dispshopdiv = NULL," & VbCRLF
		end if
		
		strSql = strSql & " doc_kind = '"&doc_kind&"'" & VbCRLF		
		strSql = strSql & " WHERE" & VbCRLF
		strSql = strSql & " doc_idx = " & iDoc_Idx & ""
		
		'response.write strSql &"<br>"
		dbget.execute strSql

		'//매장등록
	    strSql = " delete from db_shop.dbo.tbl_offshop_board_shop"&VbCRLF
        strSql = strSql & " where doc_idx="&iDoc_Idx&""
        
        'response.write strSql &"<br>"
        dbget.Execute strSql

		if dispshopidon = "ON" then
			if shopid <> "" then	
				shopid = split(shopid,",")

				if isarray(shopid) then
					for i = 0 to ubound(shopid)
					
			        strSql = " insert into db_shop.dbo.tbl_offshop_board_shop"&VbCRLF
			        strSql = strSql & " (doc_idx,shopid) values"&VbCRLF
			        strSql = strSql & " ("& iDoc_Idx &",'"& requestCheckVar(trim(shopid(i)),32) &"')"
			        
					'response.write strSql &"<br>"
					dbget.execute strSql
					
					next
				end if
		    end if
    	end if
    	        
	End If

	'####### 첨부파일 저장 #######
	If sDoc_File <> "" Then
		strSql = ""
		If iDoc_Idx <> "" Then
			strSql = " DELETE db_shop.dbo.tbl_offshop_board_file WHERE doc_idx = '" & iDoc_Idx & "' "
		End If
		vFileTemp = Split(sDoc_File, ",")
		For i = 0 To UBOUND(vFileTemp)
			strSql = strSql & " INSERT INTO db_shop.dbo.tbl_offshop_board_file " & _
							  "		(file_name, doc_idx) " & _
							  "	VALUES " & _
							  "		('" & vFileTemp(i) & "', '" & iDoc_Idx & "') " & vbCrLf
		Next
		'response.write strSql &"<br>"
		dbget.execute strSql
	Else
		If requestCheckVar(Request("isfile"),1) = "o" Then
			dbget.execute " DELETE db_shop.dbo.tbl_offshop_board_file WHERE doc_idx = '" & iDoc_Idx & "' "
		End If
	End If
	
	If Request("gubun") = "write" Then
		session.codePage = 949
		dbget.close
		Response.Write "<script type='text/javascript'>alert('OK'); location.href='offshop_board.asp?menupos="&g_MenuPos&"';</script>"
		response.end
	Else
		session.codePage = 949
		dbget.close
		Response.Write "<script type='text/javascript'>alert('OK'); location.href='offshop_board.asp?menupos="&g_MenuPos&"';</script>"
		response.end
	End If

elseif mode = "view" then

	strSql = " UPDATE db_shop.dbo.tbl_offshop_board_document SET " & _
			 " doc_status = '" & sDoc_Status & "'" & _				 
			 " WHERE " & _
			 " doc_idx = '" & iDoc_Idx & "' "
	
	'response.write strSql &"<br>"
	dbget.execute strSql

	session.codePage = 949
	dbget.close
	Response.Write "<script type='text/javascript'>alert('OK'); location.href='offshop_board.asp?menupos="&g_MenuPos&"';</script>"
	response.end
	
elseif mode = "del" then

	strSql = " UPDATE db_shop.dbo.tbl_offshop_board_document SET " & _
			 " doc_useyn = 'N'" & _				 
			 " WHERE " & _
			 " doc_idx = '" & iDoc_Idx & "' "
	
	'response.write strSql &"<br>"
	dbget.execute strSql
	
	session.codePage = 949
	dbget.close
	Response.Write "<script type='text/javascript'>alert('OK'); location.href='offshop_board.asp?menupos="&g_MenuPos&"';</script>"
	response.end	
end if

session.codePage = 949
%>

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
