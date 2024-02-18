<%@ language = vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  재고파악디비등록및 수정
' History : 2007.07.13 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->

<% 
dim mode , order , errstock , basicstock , jaego , error , idx , sql,itemid,itemoption	'변수선언
	mode = request("mode")				'받아온모드
	order = now()						'날짜
	errstock = request("errstock")		'오차
	jaego = request("jaego")			'실재파악한재고
			if jaego = "" then			'실제파악한재고 값이 업으면
				jaego = 0				'기본값 0
			end if
	
	idx = request("idx") 				'인덱스값을 받아온다.
	itemid = request("itemid")
	itemoption = request("itemoption")

'	sql = "select basicstock from [db_summary].[dbo].tbl_req_realstock"
'	sql = sql & " where 1=1 and itemid = '" & itemid & "' and itemoption = '" & itemoption & "'"
    
    sql = "select (realstock + ipkumdiv5+ offconfirmno) as basicstock from [db_summary].[dbo].tbl_current_logisstock_summary"
    sql = sql & " where itemgubun='10' and itemid = '" & itemid & "' and itemoption = '" & itemoption & "'"

	rsget.open sql,dbget,1
	
	if not rsget.eof then
		basicstock = rsget("basicstock")
	end if
	
	sql = ""
	rsget.close
	
	if basicstock = "" then
	basicstock = 0
	end if
	error = jaego - basicstock			'오차는 실제파악한재고에서 재고파악용재고를 뺀다.


%>

<!--수정모드시작-->
<%
if mode = "edit" then
	dim sql50
		sql50 = "update [db_summary].[dbo].tbl_req_realstock set realstock = "& jaego &" , errstock = "& error &""	& VbCrlf
		sql50 = sql50 & " where idx = " & idx 
		'response.write sql50
		dbget.execute sql50
	%>
	<script language="javascript">
	opener.location.reload();
	self.close();
	</script>
<!--수정모드끝-->

<!--재고파악한수량 등록모드시작-->
<% 
elseif mode = "" then
 
	dim sql12 , sql2
	sql12 = "select * from [db_summary].[dbo].tbl_req_realstock" 
	sql12 = sql12 & " where itemid = '"& itemid &"' and itemoption = '" & itemoption & "'" 
	sql12 = sql12 & " order by statecd asc"
	'response.write sql12&"<br>"
	rsget.open sql12,dbget,1
		
	if not rsget.eof then				'레코드가 있다면		
	
		sql2 = "update [db_summary].[dbo].tbl_req_realstock set"	& VbCrlf
		sql2 = sql2 & " errstock = "& error &" , actiondate = '"& order &"', realstock = "& jaego &" , basicstock = "& basicstock &", statecd = 5" 	& VbCrlf
		sql2 = sql2 & " where 1=1 and itemid = '" & itemid & "' and itemoption = '" & itemoption & "'" 
		''response.write sql2     '삑살시 뿌려본다.
		dbget.execute sql2		
	else
%>
<script language="javascript">
	if (confirm('재고파악 지시사항이 없습니다. 재고파악 지시 하시겠습니까?')){
	location.href = '/admin/stock/jaegoadd.asp?itemid=<%=itemid%>&submitview=yes';
	}
</script>	
<%	
	end if
	rsget.close

%>

	<script language="javascript">
	opener.location.reload();
	self.close();
	</script>
<% end if %>
<!--재고파악한수량 등록모드끝-->

<!-- #include virtual="/lib/db/dbclose.asp" -->