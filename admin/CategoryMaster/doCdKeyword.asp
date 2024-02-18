<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<%
'###############################################
' PageName : doCdKeyword.asp
' Discription : 카테고리 키워드 수정처리 페이지
' History : 2008.03.20 허진원 : 이전 Admin에서 이전/수정
'###############################################

dim cdl,cdm,cds,keyword

cdl = request("cdl")
cdm = request("cdm")
cds = request("cds")
keyword = trim(html2db(request("keyword")))

dim sqlStr

if cdl <> "" and cdm <> "" and cds <> "" then

	sqlStr = "update [db_item].dbo.tbl_Cate_small"
	sqlStr = sqlStr + " set keyword ='" + keyword + "'"
	sqlStr = sqlStr + " where code_large='" + cdl + "'"
	sqlStr = sqlStr + " and code_mid='" + cdm + "'"
	sqlStr = sqlStr + " and code_small='" + cds + "'"
	rsget.Open sqlStr,dbget,1

else
response.write "<script language='javascript'>alert('카테고리 데어터가 없습니다.\n카테고리를 다시 선택하고 키워드를 입력해주세요');history.back();</script>"
dbget.close()	:	response.End
end if

dim refer
refer = request.ServerVariables("HTTP_REFERER")
%>
<script language="javascript">
alert('적용 되었습니다.');
location.replace('<%= refer %>');
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->