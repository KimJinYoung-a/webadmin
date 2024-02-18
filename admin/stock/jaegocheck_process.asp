<%@ language = vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  재고파악
' History : 2007.07.13 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stockclass/jaegostock.asp"-->

<%
dim fnow,idx, fmode , jisiname , order , smallimage,itemid,makerid,itemname,itemoption,imagesrc,realstock	'변수선언
dim errstock,actionstartdate,itemgubun,jisiid,stats			'변수선언
	idx = html2db(request("idx"))							'테이블의 인덱스값을 받아온다
	fmode = html2db(request("mode")	)						'모드구분
	jisiname = html2db(request("jisiname"))					'작업지시자이름
	jisiid = html2db(session("ssBctId"))					'지시자 값을 세션 id로 받아온다.
	order = now()											'작업지시일
	smallimage = html2db(request("smallimage"))				'이미지
	itemid = html2db(request("itemid"))						'상품id
	makerid = html2db(request("makerid"))					'브랜드id
	itemname = html2db(request("itemname"))					'상품명
	itemoption = html2db(request("itemoption"))				'상품옵션코드
		if itemoption = "" then								'상품옵션코드가 업으면
			itemoption = "0000"								'기본값0 입력
		end if
	realstock = html2db(request("realstock"))				'재고파악한재고
	errstock = html2db(request("errstock"))					'오차
	actionstartdate = Left(request("actionstartdate"),10)	'재고파악일시
	itemgubun = html2db(request("itemgubun"))				'상품온,오프구분
	stats = 1												'상태 기본값 1
	imagesrc = request("imagesrc")							'제품이미지
%>
	
<% 
dim sql , refer , sql111			'변수선언

if fmode = "" then					'재고파악 지시 모드
	%>

	<%
	dim sql12
		sql12 = "select * from [db_summary].[dbo].tbl_req_realstock" 
		sql12 = sql12 & " where itemid = '"& itemid &"' order by statecd asc"
		rsget.open sql12,dbget,1
	
	if not rsget.eof then				'레코드가 있다면
		if rsget("statecd") = 1 then	'상품에서 상태값이 작업지시중(1) 이라면	
			rsget.close
	%>		
		<script language="javascript">
			alert('동일한 상품이 재고파악중입니다. 확인하신후 다시입력하세요');
			opener.location.reload();
			self.close();
			</script>
	<%
		dbget.close()	:	response.End	
		end if
	end if
	rsget.close
	
	sql = "INSERT INTO [db_summary].[dbo].tbl_req_realstock" 		'상품코드로 옵션을 쿼리해서 저장 한다.
	sql = sql & " (itemgubun,itemid,itemoption)"
	sql = sql & " select a.itemgubun,a.itemid,isnull(b.itemoption,'0000')"
	sql = sql & " from [db_item].[dbo].tbl_item a"
	sql = sql & " left join [db_item].[dbo].tbl_item_option b" 
	sql = sql & " on a.itemid = b.itemid"
	sql = sql & " where a.itemid = '" & itemid &"'"
	'response.write sql			'오류시 화면에 뿌려본다
	dbget.execute sql
	
	sql = ""
	sql = "update [db_summary].[dbo].tbl_req_realstock set"
	sql = sql & " itemgubun='" & itemgubun & "'"		& VbCrlf
	sql = sql & " ,reguserid='" & jisiid & "'"			& VbCrlf
	sql = sql & " ,statecd='" & stats & "'"
	sql = sql & " where 1=1 and itemid = '" & itemid & "' and statecd is null" 
	'response.write sql			'오류시 화면에 뿌려본다
	dbget.execute sql

	%>
					
	<script language="javascript">
		opener.location.reload();
		self.close();
	</script>	

<!--삭제모드 시작-->
	<% 
	elseif fmode = "del" then				
	sql = "delete from [db_summary].[dbo].tbl_req_realstock where idx=" & idx
	'response.write sql			'오류시 화면에 뿌려본다 	
	dbget.execute sql
	refer = request.ServerVariables("HTTP_REFERER")			'이전페이지의 내용을 가져온다
	%>
	<script language="javascript">
	location.replace('<%= refer %>');
	</script>
<!--삭제모드 끝-->


<!--반영모드시작-->
	<% elseif fmode = "banyoung" then 
	
'	response.write "수정중..."
'	dbget.close()	:	response.End
	
	sql = "exec db_summary.dbo.ten_realchekErr_Input '"& actionstartdate &"', '"& itemgubun &"', '"& itemid &"' , '"& itemoption &"', "& errstock &", '"& jisiid &"'"
	dbget.execute sql
	'response.write sql			'삑살시 뿌려본다
	
	sql111 = "update [db_summary].[dbo].tbl_req_realstock set finishdate = '"& order &"' , statecd = '7' , finishuserid = '"& jisiid &"'" 	& VbCrlf
	sql111 = sql111 & " where idx = '"& idx &"'"
	dbget.execute sql111
	'response.write sql			'삑살시 뿌려본다
	
	refer = request.ServerVariables("HTTP_REFERER")			'이전페이지의 내용을 가져온다
	%>
	
	<script language="javascript">
	{
		alert('저장되었습니다.전시여부,판매여부,사용여부,단종여부,한정여부를 수정하십시오');
		location.replace('<%= refer %>');		
		var edit = window.open("itemcurrentstock.asp?itemgubun=<%=itemgubun%>&itemid=<%=itemid%>&itemoption=<%=itemoption%>", "jaegoadd" , 'width=1024,height=768,scrollbars=yes,resizable=yes');
		edit.focus();
	}		
	</script>
	
<!--반영모드끝-->		

<% end if %>
<!-- #include virtual="/lib/db/dbclose.asp" -->