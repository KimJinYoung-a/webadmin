<%@ language = vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  브랜드별 대규모 저장
' History : 2007.07.31 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stockclass/jaegostock.asp"-->

<%
dim drawitemid , drawitemid_re
drawitemid = request("drawitemid")			'선택한 상품id를 전부 받아온다.
drawitemid_re = split(drawitemid,",")		' 콤마를 기준으로 2차원배열로 지정한다.
dim lens
lens = UBound(drawitemid_re)		'배열의 수를 체크한다.

dim fnow, fmode , order		'변수선언
dim itemgubun,jisiid,stats			'변수선언
	fmode = html2db(request("mode")	)						'모드구분
	jisiid = html2db(session("ssBctId"))					'지시자 값을 세션 id로 받아온다.
	order = now()											'작업지시일
	itemgubun = "10"				'상품온,오프구분
	stats = 1												'상태 기본값 1

%>


<%
dim oip,sql, i,sql12
dim a
for a = 0 to lens -1 		
%>	
	<% 
	sql12 = "select * from [db_summary].[dbo].tbl_req_realstock" 
	sql12 = sql12 & " where itemid = '"& drawitemid_re(a) &"' order by statecd asc"
	rsget.open sql12,dbget,1
		
	if not rsget.eof then				'레코드가 있다면
		if rsget("statecd") = 1 then	'상품에서 상태값이 작업지시중(1) 이라면	
			rsget.close
		%>		
		<script language="javascript">
			alert('상품번호(<%=drawitemid_re(a)%>) 이미 재고파악중입니다. 이전 선택 상품 등록완료! <%=drawitemid_re(a)%> 다음상품부터 다시 등록하세요');
			opener.opener.location.reload();
			opener.frm.drawitemid.value = '';		
			self.close();
		</script>
		<%
		dbget.close()	:	response.End	
		end if
	end if
	rsget.close
	%>
	
<%
		set oip = new Cfitemlist        	'클래스 지정
		oip.Frectitemid = drawitemid_re(a)		'상품별로 루프를 돌면서 sql쿼리에 상품id를 넣고 상품정보를 받아온다.
		oip.fbrandinsert()					'클래스를 실행

	for i=0 to oip.FTotalCount - 1 	' 쿼리해서 받아온 상품정보를 옵션 유,무에 따라서 뿌린다. 

	sql = "INSERT INTO [db_summary].[dbo].tbl_req_realstock (itemgubun,itemid,itemoption,reguserid,statecd) VALUES "	& VbCrlf
	sql = sql & "('" & itemgubun & "'"		& VbCrlf
	sql = sql & ",'" & drawitemid_re(a) & "'"			& VbCrlf
	sql = sql & ",'" & oip.flist(i).fitemoption & "'"		& VbCrlf
	sql = sql & ",'" & jisiid & "'"			& VbCrlf
	sql = sql & ",'" & stats & "')"
	'response.write sql&"<br>"			'오류시 화면에 뿌려본다
	dbget.execute sql
	next	
next
%>

<!-- #include virtual="/lib/db/dbclose.asp" -->

<script language="javascript">
	alert('저장되었습니다');
	opener.opener.location.reload();
	opener.frm.drawitemid.value = '';
	self.close();
</script>