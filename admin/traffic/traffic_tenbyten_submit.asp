<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/traffic/traffic_class.asp"-->
<%
'###########################################################
' Description :  텐바이텐 traffic analysis(다음에서 텐바이텐 db에저장 페이지)  
' History : 2007.09.04 한용민 생성
'###########################################################

dim ColumnValue_12,ColumnValue_13,ColumnValue_15,ColumnValue1_10,ColumnValue1_11,ColumnValue2_9
dim ColumnValue_17,ColumnValue_18,ColumnValue_20,ColumnValue1_14,ColumnValue1_15,ColumnValue2_12
dim ColumnValue_22,ColumnValue_23,ColumnValue_25,ColumnValue1_18,ColumnValue1_19,ColumnValue2_15
dim ColumnValue_27,ColumnValue_28,ColumnValue_30,ColumnValue1_22,ColumnValue1_23,ColumnValue2_18
dim ColumnValue_32,ColumnValue_33,ColumnValue_35,ColumnValue1_26,ColumnValue1_27,ColumnValue2_21
dim ColumnValue_37,ColumnValue_38,ColumnValue_40,ColumnValue1_30,ColumnValue1_31,ColumnValue2_24
dim ColumnValue_42,ColumnValue_43,ColumnValue_45,ColumnValue1_34,ColumnValue1_35,ColumnValue2_27

ColumnValue_12=right(request("ColumnValue_12"),8)
ColumnValue_13=request("ColumnValue_13")
ColumnValue_15=request("ColumnValue_15")
ColumnValue1_10=request("ColumnValue1_10")
ColumnValue1_11=request("ColumnValue1_11")
ColumnValue2_9=request("ColumnValue2_9")

ColumnValue_17=right(request("ColumnValue_17"),8)
ColumnValue_18=request("ColumnValue_18")
ColumnValue_20=request("ColumnValue_20")
ColumnValue1_14=request("ColumnValue1_14")
ColumnValue1_15=request("ColumnValue1_15")
ColumnValue2_12=request("ColumnValue2_12")

ColumnValue_22=right(request("ColumnValue_22"),8)
ColumnValue_23=request("ColumnValue_23")
ColumnValue_25=request("ColumnValue_25")
ColumnValue1_18=request("ColumnValue1_18")
ColumnValue1_19=request("ColumnValue1_19")
ColumnValue2_15=request("ColumnValue2_15")

ColumnValue_27=right(request("ColumnValue_27"),8)
ColumnValue_28=request("ColumnValue_28")
ColumnValue_30=request("ColumnValue_30")
ColumnValue1_22=request("ColumnValue1_22")
ColumnValue1_23=request("ColumnValue1_23")
ColumnValue2_18=request("ColumnValue2_18")

ColumnValue_32=right(request("ColumnValue_32"),8)
ColumnValue_33=request("ColumnValue_33")
ColumnValue_35=request("ColumnValue_35")
ColumnValue1_26=request("ColumnValue1_26")
ColumnValue1_27=request("ColumnValue1_27")
ColumnValue2_21=request("ColumnValue2_21")

ColumnValue_37=right(request("ColumnValue_37"),8)
ColumnValue_38=request("ColumnValue_38")
ColumnValue_40=request("ColumnValue_40")
ColumnValue1_30=request("ColumnValue1_30")
ColumnValue1_31=request("ColumnValue1_31")
ColumnValue2_24=request("ColumnValue2_24")

ColumnValue_42=right(request("ColumnValue_42"),8)
ColumnValue_43=request("ColumnValue_43")
ColumnValue_45=request("ColumnValue_45")
ColumnValue1_34=request("ColumnValue1_34")
ColumnValue1_35=request("ColumnValue1_35")
ColumnValue2_27=request("ColumnValue2_27")
%>

<%
dim i 
dim sqlselect1,sqlselect2,sqlselect3,sqlselect4,sqlselect5,sqlselect6,sqlselect7
dim sqlupdate1,sqlupdate2,sqlupdate3,sqlupdate4,sqlupdate5,sqlupdate6,sqlupdate7
dim yyyymmdd1,yyyymmdd2,yyyymmdd3,yyyymmdd4,yyyymmdd5,yyyymmdd6,yyyymmdd7
'##################################################################			첫번째라인
	sqlselect1 = "select yyyymmdd from db_datamart.dbo.tbl_traffic_analysis where yyyymmdd = '"& ColumnValue_12 &"'"
	'response.write sqlselect1     '삑살시 뿌려본다.
	db3_rsget.open sqlselect1,db3_dbget,1
		if not db3_rsget.eof then						'레코드의 첫번째가 아니라면
			do until db3_rsget.eof						'레코드의 끝까지 루프 ㄱㄱ
				yyyymmdd1 = db3_rsget("yyyymmdd")
				db3_rsget.movenext
			loop		
		end if
	db3_rsget.close

if yyyymmdd1 = "" then
	sqlupdate1 = "insert into db_datamart.dbo.tbl_traffic_analysis (yyyymmdd,pageview,totalcount,newcount,recount,realcount) values"	& VbCrlf
	sqlupdate1 = sqlupdate1 & " ("& ColumnValue_12 &","& ColumnValue_13 &","& ColumnValue_15 &","& ColumnValue1_10 &","& ColumnValue1_11 &","& ColumnValue2_9 &")" 	
	'response.write sqlupdate1     '삑살시 뿌려본다.
	db3_dbget.execute sqlupdate1
end if	
%>
<%
'##################################################################			두번째라인
sqlselect2 = "select yyyymmdd from db_datamart.dbo.tbl_traffic_analysis where yyyymmdd = '"& ColumnValue_17 &"'"
	'response.write sqlselect2     '삑살시 뿌려본다.
	db3_rsget.open sqlselect2,db3_dbget,1
		if not db3_rsget.eof then						'레코드의 첫번째가 아니라면
			do until db3_rsget.eof						'레코드의 끝까지 루프 ㄱㄱ
				yyyymmdd2 = db3_rsget("yyyymmdd")
				db3_rsget.movenext
			loop		
		end if
	db3_rsget.close

if yyyymmdd2 = "" then
	sqlupdate2 = "insert into db_datamart.dbo.tbl_traffic_analysis (yyyymmdd,pageview,totalcount,newcount,recount,realcount) values"	& VbCrlf
	sqlupdate2 = sqlupdate2 & " ("& ColumnValue_17 &","& ColumnValue_18 &","& ColumnValue_20 &","& ColumnValue1_14 &","& ColumnValue1_15 &","& ColumnValue2_12 &")" 	
	'response.write sqlupdate2     '삑살시 뿌려본다.
	db3_dbget.execute sqlupdate2
end if	
%>
<%
'##################################################################			3번째라인
sqlselect3 = "select yyyymmdd from db_datamart.dbo.tbl_traffic_analysis where yyyymmdd = '"& ColumnValue_22 &"'"
	'response.write sqlselect3     '삑살시 뿌려본다.
	db3_rsget.open sqlselect3,db3_dbget,1
		if not db3_rsget.eof then						'레코드의 첫번째가 아니라면
			do until db3_rsget.eof						'레코드의 끝까지 루프 ㄱㄱ
				yyyymmdd3 = db3_rsget("yyyymmdd")
				db3_rsget.movenext
			loop		
		end if
	db3_rsget.close

if yyyymmdd3 = "" then
	sqlupdate3 = "insert into db_datamart.dbo.tbl_traffic_analysis (yyyymmdd,pageview,totalcount,newcount,recount,realcount) values"	& VbCrlf
	sqlupdate3 = sqlupdate3 & " ("& ColumnValue_22 &","& ColumnValue_23 &","& ColumnValue_25 &","& ColumnValue1_18 &","& ColumnValue1_19 &","& ColumnValue2_15 &")" 	
	'response.write sqlupdate3     '삑살시 뿌려본다.
	db3_dbget.execute sqlupdate3
end if	
%>
<%
'##################################################################			4번째라인
sqlselect4 = "select yyyymmdd from db_datamart.dbo.tbl_traffic_analysis where yyyymmdd = '"& ColumnValue_27 &"'"
'	response.write sqlselect4     '삑살시 뿌려본다.
	db3_rsget.open sqlselect4,db3_dbget,1
		if not db3_rsget.eof then						'레코드의 첫번째가 아니라면
			do until db3_rsget.eof						'레코드의 끝까지 루프 ㄱㄱ
				yyyymmdd4 = db3_rsget("yyyymmdd")
				db3_rsget.movenext
			loop		
		end if
	db3_rsget.close

if yyyymmdd4 = "" then
	sqlupdate4 = "insert into db_datamart.dbo.tbl_traffic_analysis (yyyymmdd,pageview,totalcount,newcount,recount,realcount) values"	& VbCrlf
	sqlupdate4 = sqlupdate4 & " ("& ColumnValue_27 &","& ColumnValue_28 &","& ColumnValue_30 &","& ColumnValue1_22 &","& ColumnValue1_23 &","& ColumnValue2_18 &")" 	
	'response.write sqlupdate4     '삑살시 뿌려본다.
	db3_dbget.execute sqlupdate4
end if	
%>
<%
'##################################################################			5번째라인
sqlselect5 = "select yyyymmdd from db_datamart.dbo.tbl_traffic_analysis where yyyymmdd = '"& ColumnValue_32 &"'"
	'response.write sqlselect5     '삑살시 뿌려본다.
	db3_rsget.open sqlselect5,db3_dbget,1
		if not db3_rsget.eof then						'레코드의 첫번째가 아니라면
			do until db3_rsget.eof						'레코드의 끝까지 루프 ㄱㄱ
				yyyymmdd5 = db3_rsget("yyyymmdd")
				db3_rsget.movenext
			loop		
		end if
	db3_rsget.close

if yyyymmdd5 = "" then
	sqlupdate5 = "insert into db_datamart.dbo.tbl_traffic_analysis (yyyymmdd,pageview,totalcount,newcount,recount,realcount) values"	& VbCrlf
	sqlupdate5 = sqlupdate5 & " ("& ColumnValue_32 &","& ColumnValue_33 &","& ColumnValue_35 &","& ColumnValue1_26 &","& ColumnValue1_27 &","& ColumnValue2_21 &")" 	
	response.write sqlupdate5     '삑살시 뿌려본다.
	db3_dbget.execute sqlupdate5
end if	
%>
<%
'##################################################################			6번째라인
sqlselect6 = "select yyyymmdd from db_datamart.dbo.tbl_traffic_analysis where yyyymmdd = '"& ColumnValue_37 &"'"
	'response.write sqlselect6     '삑살시 뿌려본다.
	db3_rsget.open sqlselect6,db3_dbget,1
		if not db3_rsget.eof then						'레코드의 첫번째가 아니라면
			do until db3_rsget.eof						'레코드의 끝까지 루프 ㄱㄱ
				yyyymmdd6 = db3_rsget("yyyymmdd")
				db3_rsget.movenext
			loop		
		end if
	db3_rsget.close

if yyyymmdd6 = "" then
	sqlupdate6 = "insert into db_datamart.dbo.tbl_traffic_analysis (yyyymmdd,pageview,totalcount,newcount,recount,realcount) values"	& VbCrlf
	sqlupdate6 = sqlupdate6 & " ("& ColumnValue_37 &","& ColumnValue_38 &","& ColumnValue_40 &","& ColumnValue1_30 &","& ColumnValue1_31 &","& ColumnValue2_24 &")" 	
	response.write sqlupdate6     '삑살시 뿌려본다.
	db3_dbget.execute sqlupdate6
end if	
%>
<%
'##################################################################			6번째라인
sqlselect7 = "select yyyymmdd from db_datamart.dbo.tbl_traffic_analysis where yyyymmdd = '"& ColumnValue_42 &"'"
	'response.write sqlselect7     '삑살시 뿌려본다.
	db3_rsget.open sqlselect7,db3_dbget,1
		if not db3_rsget.eof then						'레코드의 첫번째가 아니라면
			do until db3_rsget.eof						'레코드의 끝까지 루프 ㄱㄱ
				yyyymmdd7 = db3_rsget("yyyymmdd")
				db3_rsget.movenext
			loop		
		end if
	db3_rsget.close

if yyyymmdd7 = "" then
	sqlupdate7 = "insert into db_datamart.dbo.tbl_traffic_analysis (yyyymmdd,pageview,totalcount,newcount,recount,realcount) values"	& VbCrlf
	sqlupdate7 = sqlupdate7 & " ("& ColumnValue_42 &","& ColumnValue_43 &","& ColumnValue_45 &","& ColumnValue1_34 &","& ColumnValue1_31 &","& ColumnValue2_27 &")" 	
	response.write sqlupdate7     '삑살시 뿌려본다.
	db3_dbget.execute sqlupdate7
end if	
%>
<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db3close.asp" -->
	<script language="javascript">
	opener.location.reload();
	self.close();
	</script>
	