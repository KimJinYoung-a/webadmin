<%@ language=vbscript %>
<% option explicit %>
<%
'###############################################
' PageName : doRelate_Process.asp
' Discription : 카테고리 관련 키워드 처리 페이지
' History : 2008.03.29 허진원 생성
'			2022.07.05 한용민 수정(isms취약점조치, 표준코딩으로변경)
'###############################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
dim mode, linkCode, cdl,cdm,cds, linkKeyword, linkURL, menupos
	menupos = requestcheckvar(getNumeric(request("menupos")),10)
	mode = requestcheckvar(request("mode"),32)
	LinkCode = requestcheckvar(getNumeric(request("rid")),10)
cdl			= request("cdl")
cdm			= request("cdm")
cds			= request("cds")
linkKeyword = trim(html2db(request("linkKeyword")))
linkURL		= trim(html2db(request("linkURL")))

dim sqlStr

'// 처리 분기
Select Case mode
	Case "add"
		if linkKeyword <> "" and not(isnull(linkKeyword)) then
			linkKeyword = ReplaceBracket(linkKeyword)
		end If
		if linkURL <> "" and not(isnull(linkURL)) then
			linkURL = ReplaceBracket(linkURL)
		end If

		'### 신규 등록 ###
		sqlStr = "insert into [db_item].dbo.tbl_Cate_RelateLink"
		sqlStr = sqlStr + " (code_large, code_mid, code_small, linkKeyword, linkURL) values "
		sqlStr = sqlStr + " ('" + cdl + "'"
		sqlStr = sqlStr + " ,'" + cdm + "'"
		sqlStr = sqlStr + " ,'" + cds + "'"
		sqlStr = sqlStr + " ,'" + linkKeyword + "'"
		sqlStr = sqlStr + " ,'" + linkURL + "')"
		rsget.Open sqlStr,dbget,1

	Case "modify"
		if linkKeyword <> "" and not(isnull(linkKeyword)) then
			linkKeyword = ReplaceBracket(linkKeyword)
		end If
		if linkURL <> "" and not(isnull(linkURL)) then
			linkURL = ReplaceBracket(linkURL)
		end If

		'### 수정 ###
		sqlStr = "update [db_item].dbo.tbl_Cate_RelateLink"
		sqlStr = sqlStr + " set code_large ='" + cdl + "'"
		sqlStr = sqlStr + "		,code_mid ='" + cdm + "'"
		sqlStr = sqlStr + "		,code_small ='" + cds + "'"
		sqlStr = sqlStr + "		,linkKeyword ='" + linkKeyword + "'"
		sqlStr = sqlStr + "		,linkURL ='" + linkURL + "'"
		sqlStr = sqlStr + " where linkCode='" + linkCode + "'"
		rsget.Open sqlStr,dbget,1

	Case "delete"
		'### 삭제 ###
		sqlStr = "delete [db_item].dbo.tbl_Cate_RelateLink"
		sqlStr = sqlStr + " where linkCode='" + linkCode + "'"
		rsget.Open sqlStr,dbget,1

	Case Else
		response.write "<script type='text/javascript'>alert('정상적인 접근이 아닙니다.\n관리자에게 연락주세요.');history.back();</script>"
		dbget.close()	:	response.End
end Select

%>
<script type='text/javascript'>
alert('적용 되었습니다.');
location.replace("RelateKeywordLink_list.asp?cdl=<%=cdl%>&cdm=<%=cdm%>&cds=<%=cds%>&menupos=<%= menupos %>");
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->