<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  월간원가보고서
' History : 2007.09.10 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db3open.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/wonga/wonga_month_class.asp"-->

<%
dim gubun,yyyy,mm
gubun = Request("gubunbox")			'그룹 검색에 필요한변수
yyyy = request("yyyybox")		'년도검색에 필요한변수
mm = request("mmbox")			'달 검색에 필요한변수

dim owongamonth,i
set owongamonth = new Cwongalist
	owongamonth.frectyyyy = request("yyyybox")
	owongamonth.frectmm = request("mmbox")
	owongamonth.frectgubun = Request("gubunbox")
	owongamonth.fwongamonth()

dim owongamonth_re
set owongamonth_re = new Cwongalist
	owongamonth_re.frectgubun = Request("gubunbox")
	owongamonth_re.fwongamonth_add()
		
dim ocwongalist 
set ocwongalist = new Cwongalist
	ocwongalist.frectyyyy = request("yyyybox")
	ocwongalist.frectgubun = Request("gubunbox")
	ocwongalist.fwongalist()
%>	
<%	
'Const adOpenKeyset = 1
'Const adLockReadOnly = 1
'Const adUseClient = 3
'########################################################### 구분셀렉트박스	
Sub DrawUserGubun(gubunbox,gubunid)		'검색하고자하는 것을 셀렉트 박스네임에 넣고, 디비에 있는 값을 검색._selectboxname은 sub구문에서만 쓰임
	dim userquery, tem_str
	
	'사용자 검색 옵션 내용 DB에서 가져오기
	userquery = "select groupname from"
	userquery = userquery & " db_datamart.dbo.tbl_month_wonga"
	userquery = userquery & " group by groupname"
	userquery = userquery & " order by groupname asc"
	db3_rsget.Open userquery, db3_dbget, 1

	response.write "<select name='" & gubunbox & "'>"		'검색하고자하는 것을 셀렉트 네임으로 하고
	response.write "<option value=''"							'옵션의 값이 없으면
		if gubunid ="" then									'디비에서 검색할 값이 없으므로,
			response.write "selected"
		end if
	response.write ">선택</option>"								'선택이란 단어가 나오도록.

	if not db3_rsget.EOF then
		do until db3_rsget.EOF
			if Lcase(gubunid) = Lcase(db3_rsget("groupname")) then 	'검색될 이름과 db에 저장된 이름을 비교해서 맞다면, //
				tem_str = " selected"								'// 검색어로 선택
			end if
			response.write "<option value='" & db3_rsget("groupname") & "' " & tem_str & ">" & db2html(db3_rsget("groupname")) & "</option>"
			tem_str = ""				'db3_rsget에 gubunid 선택하고 검색할 값으로 선택
			db3_rsget.movenext
		loop
	end if
	response.write "</select> &nbsp; 년도 : "
db3_rsget.close		
End Sub	
'########################################################### 년도 셀렉트박스
Sub DrawyyyyGubun(yyyybox,yyyyid)		'검색하고자하는 것을 셀렉트 박스네임에 넣고, 디비에 있는 값을 검색._selectboxname은 sub구문에서만 쓰임
	dim userquery, tem_str
	
	'사용자 검색 옵션 내용 DB에서 가져오기
	userquery = "select left(yyyymm,4) as yyyymm from"
	userquery = userquery & " db_datamart.dbo.tbl_month_wonga"
	userquery = userquery & " group by left(yyyymm,4)"
	userquery = userquery & " order by left(yyyymm,4) asc"
	db3_rsget.Open userquery, db3_dbget, 1

	response.write "<select name='" & yyyybox & "'>"		'검색하고자하는 것을 셀렉트 네임으로 하고
	response.write "<option value=''"							'옵션의 값이 없으면
		if yyyyid ="" then									'디비에서 검색할 값이 없으므로,
			response.write "selected"
		end if
	response.write ">선택</option>"								'선택이란 단어가 나오도록.
	
	'db3_rsget.movefirst
	if not db3_rsget.EOF then
		
		do until db3_rsget.EOF
			if Lcase(yyyyid) = Lcase(left(db3_rsget("yyyymm"),4)) then 	'검색될 이름과 db에 저장된 이름을 비교해서 맞다면, //
				tem_str = " selected"								'// 검색어로 선택
			end if
			response.write "<option value='" & left(db3_rsget("yyyymm"),4) & "' " & tem_str & ">" & db2html(left(db3_rsget("yyyymm"),4)) & "</option>"
			tem_str = ""				'db3_rsget에 yyyyid 선택하고 검색할 값으로 선택
			db3_rsget.movenext
		loop
	end if
	response.write "</select> &nbsp; 달 : "
	db3_rsget.close	
End Sub		
	'########################################################### 달 셀렉트박스
Sub DrawmmGubun(mmbox,mmid)		'검색하고자하는 것을 셀렉트 박스네임에 넣고, 디비에 있는 값을 검색._selectboxname은 sub구문에서만 쓰임
	dim userquery, tem_str
	
	'사용자 검색 옵션 내용 DB에서 가져오기
	userquery = "select right(yyyymm,2) as yyyymm from"
	userquery = userquery & " db_datamart.dbo.tbl_month_wonga"
	userquery = userquery & " group by right(yyyymm,2)"
	userquery = userquery & " order by right(yyyymm,2) asc"
	db3_rsget.Open userquery, db3_dbget, 1
	response.write "<select name='" & mmbox & "'>"		'검색하고자하는 것을 셀렉트 네임으로 하고
	response.write "<option value=''"							'옵션의 값이 없으면
		if mmid ="" then									'디비에서 검색할 값이 없으므로,
			response.write "selected"
		end if
	response.write ">선택</option>"								'선택이란 단어가 나오도록.
	
	'db3_rsget.movefirst
	if not db3_rsget.EOF then
		
		do until db3_rsget.EOF
			if Lcase(mmid) = Lcase(right(db3_rsget("yyyymm"),2)) then 	'검색될 이름과 db에 저장된 이름을 비교해서 맞다면, //
				tem_str = " selected"								'// 검색어로 선택
			end if
			response.write "<option value='" & right(db3_rsget("yyyymm"),2) & "' " & tem_str & ">" & db2html(right(db3_rsget("yyyymm"),2)) & "</option>"
			tem_str = ""				'db3_rsget에 mmid 선택하고 검색할 값으로 선택
			db3_rsget.movenext
		loop
	end if
	response.write "</select>"
db3_rsget.close	
End Sub
	'###########################################################
%>

<script lagnuage="javascript">
function reg(menupos){
	var popwin = window.open('/admin/wonga/wonga_add.asp?menupos='+menupos,'reg','width=1024,height=768,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function del(gubun,aa){
var a
	a = confirm('그룹과 해당하는 카테고리&데이터 모두를 삭제를 하시겠습니까?');
	if(a==true){
		var popwin = window.open('/admin/wonga/wonga_del.asp?groupname='+gubun+'&mode='+aa,'del','width=1024,height=768,scrollbars=yes,resizable=yes');
		popwin.focus();		
	}
}
	
function del1(gubun,yyyymm){
var a
	a = confirm('그룹중 선택하신 날짜를 삭제를 하시겠습니까?');
	if(a==true){
		var popwin = window.open('/admin/wonga/wonga_del.asp?groupname='+gubun+'&yyyymm='+yyyymm,'del1','width=1024,height=768,scrollbars=yes,resizable=yes');
		popwin.focus();		
	}
}


function del2(gubun,category,field,aa){
var a
	a = confirm('정말  삭제 하시겠습니까?');
	if(a==true){
		var popwin = window.open('/admin/wonga/wonga_del.asp?groupname='+gubun+'&category='+category+'&field='+field+'&mode='+aa,'del2','width=1024,height=768,scrollbars=yes,resizable=yes');
		popwin.focus();		
	}
}


function edit(gubun,category,field,yyyymm,chulgocount){
	var popwin = window.open('/admin/wonga/wonga_edit.asp?groupname='+gubun+'&category='+category+'&field='+field+'&yyyymm='+yyyymm+'&chulgocount='+chulgocount,'edit','width=600,height=500,scrollbars=yes,resizable=yes');
	popwin.focus();	
}
	
</script>

<!--표 헤드시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
<form name="frm" method=get>
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr height="10" valign="bottom">
		<td width="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
		<td valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
		<td width="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td background="/images/tbl_blue_round_06.gif">
			<img src="/images/icon_star.gif" align="absbottom">
			<font color="red"><strong><%= gubun %> 월간 원가 보고서</strong></font>
			</td>			
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	<tr valign="top">
		<td background="/images/tbl_blue_round_04.gif"></td>
		<td><br>구분: <% DrawUserGubun "gubunbox", gubun %> &nbsp; <% DrawyyyyGubun "yyyybox",yyyy %> &nbsp; <% DrawmmGubun "mmbox",mm%>
	       	<input type=button value="검색" onclick="document.frm.submit();">
	       <p align="right"><a href="javascript:reg('<%= menupos %>');">등록하기</a></p>
	       	</td>
		<td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</form>
</table>
<!--표 헤드끝-->
<% if owongamonth.ftotalcount > 0 then %>		 <!-- 레코드 값이있다면-->
	<table width="100%" border="0" class="a" cellpadding="0" cellspacing="1" bgcolor="#BABABA" align="center">
		<tr>
			<td  bgcolor="FFFFFF" colspan="2">
				&nbsp; <%= gubun %> 비용 내역
			</td>
		</tr>
		<tr>
			<td>
				<table width="100%" border="0" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA" align="center">
					<tr bgcolor=#DDDDFF>
						<td align="center" colspan="2">구분</td>
						<td align="center"><font color="red"><%= left(owongamonth.flist(i).yyyymm,4) %>년 <%= right(owongamonth.flist(i).yyyymm,2) %>월</font></td>
						<td align="center"><font color="red">출고건당비용</font></td>
						<td align="center">기준</td>
						<td align="center">비고</td>			
					</tr>

			<% 
			dim category0_gijun_sogee,category1_gijun_sogee,category2_gijun_sogee,category3_gijun_sogee,category4_gijun_sogee,category5_gijun_sogee 
			dim gubun0_value_sum , gubun0_count_sum 
			dim gubun1_value_sum , gubun1_count_sum 
			dim gubun2_value_sum , gubun2_count_sum
			dim gubun3_value_sum , gubun3_count_sum
			dim gubun4_value_sum , gubun4_count_sum
			dim gubun5_value_sum , gubun5_count_sum
			dim gubun_value_sum , gubun_count_sum , gubun_gijun_sum
			dim gubun_value_sum_total , gubun_count_sum_total , category_gijun_sogee_total
			%>					
						
	<%
	dim sql ,ftotalcount ,t,category_idx,rectcategory_idx
		sql = "select"
		sql = sql & " category"
		sql = sql & " from db_datamart.dbo.tbl_month_wonga_category"
		sql = sql & " where 1=1 and groupname= '"& gubun &"' and category_isusing='y'"
		sql = sql & " group by category" 	
	db3_rsget.open sql,db3_dbget,1
	'response.write sql&"<br>"	
	ftotalcount = db3_rsget.recordcount
		if not db3_rsget.eof then				
			do until db3_rsget.eof				
				category_idx = category_idx&db3_rsget("category")&","				
				db3_rsget.movenext	
			loop			
		end if
	db3_rsget.close
	rectcategory_idx = left(category_idx,len(category_idx)-1)
	%>	
	
	<% dim rowspan %>
	
	<!-- 카테고리 시작 -->
	<% for t = 0 to ftotalcount - 1 %>			
		<%
		dim sql1 ,ffieldcount ,a
		sql1 = "select field"
		sql1 = sql1 & " from db_datamart.dbo.tbl_month_wonga_category"
		sql1 = sql1 & " where 1=1 and groupname= '"& gubun &"' and category_isusing='y' and category='"& t &"'"
	
		db3_rsget.open sql1,db3_dbget,1
		'response.write sql1&"<br>"	
		ffieldcount = db3_rsget.recordcount
		db3_rsget.close
		%>
		
		<!-- 필드 시작-->
			<% for a = 0 to ffieldcount -1 %>				
				<tr bgcolor=#ffffff>
					<% if rowspan = "" then %>
					<td align="center" rowspan="<%= ffieldcount %>">
						<%= frectcategoryname(t,a) %>
					</td>					
					<% end if %>	
					<td align="center">
						<%= frectfieldname(t,a) %>
					</td>
					<td align="center">
						<%= CurrFormat(frectfieldvalue(t,a)) %>
					</td>
					<td align="center">
						<%= CurrFormat(round(frectchulgovalue(t,a),0)) %>
					</td>								
					<td align="center">
						<%= CurrFormat(frectgijunvalue(t,a)) %>
					</td>
					<td align="center">
					<a href="javascript:edit('<%= gubun %>','<%= t %>','<%= a %>','<%= yyyy&mm %>','<%= owongamonth.flist(i).chulgocount %>');">수정</a>&nbsp;
					<!--<a href="javascript:del2('<%= gubun %>','<%= t %>','<%= a %>','del');">삭제</a>-->	
					</td>
				</tr>
				<% gubun0_value_sum =frectfieldvalue(t,a)+ gubun0_value_sum %>
				<% category0_gijun_sogee = cint(frectgijunvalue(t,a))+category0_gijun_sogee %>
				<% rowspan = 1 %>	
			<% next %>
				<% rowspan = "" %>		
				<tr bgcolor=#ffffff>
				<td align="center" colspan="2">
					<%= frectcategoryname(t,0) %> 소계
				</td>
				<td align="center">
					<% gubun_value_sum_total = gubun0_value_sum + gubun_value_sum_total%>
					<%= CurrFormat(gubun0_value_sum) %>
				</td>
				<td align="center">					
					<% gubun0_count_sum = gubun0_value_sum / cint(owongamonth.flist(i).chulgocount) %>
					<% gubun_count_sum_total = gubun0_count_sum + gubun_count_sum_total %>
					<%= CurrFormat(round(gubun0_count_sum)) %>
				</td>
				<td align="center">
					<% category_gijun_sogee_total = category0_gijun_sogee + category_gijun_sogee_total %>
					<%= CurrFormat(category0_gijun_sogee) %>
				</td>
				<td align="center"></td>
				<% gubun0_value_sum = 0 %><% category0_gijun_sogee = 0 %>
				</tr>
	<% next %>				
	<!-- 필드 끝 -->
			
	<!-- 카테고리 끝 -->
				
					<tr bgcolor="DDDDFF">
						<td align="center" colspan="2">
							운영비총계
						</td>
						<td align="center">
						<%= CurrFormat(gubun_value_sum_total) %>
						</td>
						<td align="center">
							<%= CurrFormat(gubun_count_sum_total) %>
						</td>
						<td align="center">
							<%= CurrFormat(category_gijun_sogee_total) %>
						</td>
						<td align="center"></td>
					</tr>
								
				</table>						
			</td>
	
	</tr>
</table>
	<br>

<% dim gubun0_sum ,gubun_totalsum ,category_sum%>
	<table width="100%" border="0" class="a" cellpadding="0" cellspacing="1" bgcolor="#BABABA" align="center">
		<tr>
			<td  bgcolor="F4F4F4" colspan=2>
				&nbsp; <%= yyyy %>년 월별 <%= gubun %>총 비용
			</td>
			<td bgcolor="ffffff" colspan=100 align="right"><input type="button" name="delbutton" value="<%=gubun%> 그룹삭제" onclick="del('<%=gubun%>','total_del');">		
			</td>
		</tr>
		<tr bgcolor=#DDDDFF>
			<td align="center">
		<!-- 년도 시작 -->
				<table width="100%" border="0" class="a" cellpadding="1" cellspacing="1" bgcolor="#BABABA" align="center">
					<tr bgcolor=#DDDDFF>
						<td align="center">년월
						</td>
					</tr>
					<% for i=0 to ocwongalist.FTotalCount - 1 %>
						<tr bgcolor=#FFFFFF>
							<td align="center"><%= ocwongalist.flist(i).yyyymm %> &nbsp;&nbsp; <a href="javascript:del1('<%=gubun%>','<%= ocwongalist.flist(i).yyyymm %>');">삭제</a>
							</td>
						</tr>
					<% next %>
					<tr bgcolor=#DDDDFF>
						<td align="center">누적통계
						</td>
					</tr>
				</table>
		<!-- 년도 끝 -->	
			</td>
			
			<% for t = 0 to ftotalcount - 1 %>
			<td align="center">
		<!-- 카테고리 시작 -->	
				<table width="100%" border="0" class="a" cellpadding="1" cellspacing="1" bgcolor="#BABABA" align="center">					
					<tr bgcolor=#DDDDFF>
						<td align="center"><%= frectcategoryname(t,0) %>
						</td>
					</tr>
					<% for i=0 to ocwongalist.FTotalCount - 1 %>
						<tr bgcolor=#FFFFFF>
							<td align="center">	
								<%= CurrFormat(frectfieldvaluesum(gubun,ocwongalist.flist(i).yyyymm,t)) %><% gubun0_sum = gubun0_sum+ frectfieldvaluesum(gubun,ocwongalist.flist(i).yyyymm,t) %>
							</td>
						</tr>							
					<% next %>
					<tr bgcolor=#DDDDFF>
						<td align="center"><%= CurrFormat(gubun0_sum) %><% gubun0_sum = 0 %></td>
					</tr>
				</table>
		<!-- 카테고리 끝 -->			
			</td>

		<% next %>
			<td align="center">
		<!-- 총운영비 시작 -->		
				<table width="100%" border="0" class="a" cellpadding="1" cellspacing="1" bgcolor="#BABABA" align="center">
					<tr bgcolor=#DDDDFF>
						<td align="center">총운영비
						</td>
					</tr>
					<% for i=0 to ocwongalist.FTotalCount - 1 %>						
						<tr bgcolor=#ffffff>
							<td align="center">
							<%= CurrFormat(frectfieldvaluesum(gubun,ocwongalist.flist(i).yyyymm,rectcategory_idx)) %>
							<% category_sum = category_sum + clng(frectfieldvaluesum(gubun,ocwongalist.flist(i).yyyymm,rectcategory_idx)) %>
							</td>
						</tr>
					<% next %>
					<tr bgcolor=#DDDDFF>
						<td align="center">
							<%= CurrFormat(category_sum) %>
						
						</td>
					</tr>	
				</table>
		<!-- 총운영비 끝 -->		
			</td>
			
		</tr>
	</table>
<% else %>
		<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA">
		    <tr align="center" bgcolor="#DDDDFF">
		    	<td align=center bgcolor="#FFFFFF"> 검색 결과가 없습니다.</td>
		    </tr>
		</table>
<% end if %>	

<!-- 표 하단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
    <tr valign="top" height="25">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td valign="bottom" align="right">&nbsp;</td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
    <tr valign="bottom" height="10">
        <td width="10" align="right"><img src="/images/tbl_blue_round_07.gif" width="10" height="10"></td>
        <td background="/images/tbl_blue_round_08.gif"></td>
        <td width="10" align="left"><img src="/images/tbl_blue_round_09.gif" width="10" height="10"></td>
    </tr>
</table>
<!-- 표 하단바 끝-->

<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/db3close.asp" -->
<!-- #include virtual="/lib/db/dbclose.asp" -->


