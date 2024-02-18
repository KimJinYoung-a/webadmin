<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  월간원가보고서 신규등록 날짜와 달 , 그룹명 디비 중복 검색페이지
' History : 2007.09.28 한용민 생성
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<%
dim yyyy,mm,groupname , yyyymm
	yyyy = request("yyyy")
	mm = request("mm")
	groupname = request("groupname")
	yyyymm = yyyy&mm
%>

<%
dim sql,i

sql = "select groupname , yyyymm" 
sql = sql & " from db_datamart.dbo.tbl_month_wonga"
sql = sql & " where 1=1 and	groupname = '"& groupname &"' and yyyymm = '"& yyyymm &"'"
'response.write sql&"<br>"
db3_rsget.open sql,db3_dbget,1

	if not db3_rsget.eof then				'레코드가 있다면

		db3_rsget.close
	%>		
		<script language="javascript">
			alert('입력하신 구분과 동일한 년도와 달이 등록되어 있습니다. 확인하신후 다시 입력하세요.');
			opener.location.reload();
			self.close();
		</script>
	<%
		dbget.close()	:	response.End
	else
		db3_rsget.close
	%>		
		<script language="javascript">
			alert('사용가능');
			self.close();
		</script>
	<% end if %>
		
	
<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->