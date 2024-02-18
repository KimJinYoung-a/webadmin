<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  재고파악리스트페이지(일반사용자용)
' History : 2007.07.13 한용민 개발
' History : 2007.11.28 한용민 수정
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stockclass/jaegostock.asp"-->

<%
dim page, pagenum , makerid , stats , orderingdate
	Page = Request("Page")					
		if Page = "" then					
			Page = 1 					
		end if
	makerid = html2db(request("makerid"))	'검색시 사용할 브랜드아디
	stats = request("stats")				'검색시 사용할 상태값
	orderingdate = request("orderingdate")	'검색시 사용할 작업지시일
%>

<%
dim oip, i
	set oip = new Cfitemlist        		
	oip.FPageSize = 15						
	oip.Fcurrpage = Page
	oip.frectmakerid = makerid
	oip.frectstats = stats
	oip.frectguestlist = "1,5" 
	oip.fjonglist()							


<!--진행상태로검색시작-->
Sub Drawstats(selectboxname, stats)		
	dim userquery, tem_str ,a

	response.write "<select name='" & selectboxname & "'>"		
	response.write "<option value=''"							
		if makerid ="" then									
			response.write "selected"
		end if
	response.write ">선택</option>"								

	'사용자 검색 옵션 내용 DB에서 가져오기
	userquery = " select statecd from [db_summary].[dbo].tbl_req_realstock"
	userquery = userquery + " where statecd=1 or statecd=3 or statecd=5 "
	userquery = userquery + " group by statecd " 'group by
	userquery = userquery + " order by statecd desc"
	rsget.Open userquery, dbget, 1

	if not rsget.EOF then
		do until rsget.EOF
			if Lcase(stats) = Lcase(rsget("statecd")) then 	
				tem_str = " selected"							
			end if
			
			if rsget("statecd") = 1 then
					a = "작업지시"
			elseif rsget("statecd") = 5 then
					a = "재고파악완료"
			elseif rsget("statecd") = 7 then
					a = "재고반영완료"
			elseif rsget("statecd") = 8 then
					a = "미반영완료" 
			end if	
			response.write "<option value='" & rsget("statecd") & "' " & tem_str & ">" & a & "</option>"
			tem_str = ""			
			rsget.movenext
		loop
	end if
	rsget.close
	response.write "</select>"
End Sub
<!--진행상태로검색끝-->
%>

<script language="javascript">
	function goSubmit(){
	frm.submit();
		}
	
	function NextPage(page){
		frm.page.value= page;
		frm.submit();
		}
	
	
	function insert(itemid,itemoption){
			var a;
			a = window.open('jaegoedit.asp?itemid='+ itemid +'&itemoption=' + itemoption,'insert','width=800,height=600,scrollbars=yes,resizable=yes');
			a.focus();
			}
	
	function edit(idx,mode){
		var edit = window.open("jaegoedit.asp?idx=" +idx + " &mode=" +mode , "edit" , 'width=600,height=600,scrollbars=yes,resizable=yes');
		edit.focus();
		}		
		
	function AnSelectAllFrame(bool){
		var frm;
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				if (frm.cksel.disabled!=true){
					frm.cksel.checked = bool;
					AnCheckClick(frm.cksel);
				}
			}
		}
	}	
		
	function AnCheckClick(e){
		if (e.checked)
			hL(e);
		else
			dL(e);
	}	
	
	function ckAll(icomp){
		var bool = icomp.checked;
		AnSelectAllFrame(bool);
	}
	
	function CheckSelected(){
		var pass=false;
		var frm;
	
		for (var i=0;i<document.forms.length;i++){
			frm = document.forms[i];
			if (frm.name.substr(0,9)=="frmBuyPrc") {
				pass = ((pass)||(frm.cksel.checked));
			}
		}
	
		if (!pass) {
			return false;
		}
		return true;
	}
	
	function print(upfrm){
	if (!CheckSelected()){
			alert('선택아이템이 없습니다.');
			return;
		}
		var frm;
			for (var i=0;i<document.forms.length;i++){
				frm = document.forms[i];
				if (frm.name.substr(0,9)=="frmBuyPrc") {
					if (frm.cksel.checked){
						upfrm.fidx.value = upfrm.fidx.value + frm.idx.value + "," ;
						
					}
				}
			}
				var tot;
				tot = upfrm.fidx.value;
			var aa;
			aa = window.open("jaegoprint.asp?idx=" +tot, "jaegoprint","width=1024,height=768,scrollbars=yes,resizable=yes");
			aa.focus();
	}
</script>

<!-- 검색 시작 -->
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="">
	<input type="hidden" name="fidx">
	<input type="hidden" name="menupos" value="<%= menupos %>">
	<tr align="center" bgcolor="#FFFFFF" >
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
		<td align="left">
        	브랜드 : <% drawSelectBoxDesignerwithName "makerid", makerid %>&nbsp;
			진행상태 : <% Drawstats "stats", stats %>&nbsp;
			<input type="hidden" name="mode"> 	
		</td>	
		<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
			<input type="button" class="button_s" value="검색" onClick="goSubmit();">
		</td>
	</tr>
	<tr align="center" bgcolor="#FFFFFF" >
		<td align="left">
		
		</td>
	</tr>
	</form>
</table>
<!-- 검색 끝 -->

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
	<tr>
		<td align="left">

		</td>
		<td align="right">	
			<input type="button" class="button" value="선택시트출력" onclick="print(frm)">			
		</td>
	</tr>
	</form>
</table>
<!-- 액션 끝 -->

<!-- 리스트 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<% if oip.FResultCount > 0 then %>
	<tr height="25" bgcolor="FFFFFF">
		<td colspan="15">
			검색결과 : <b><%= oip.FTotalCount %></b>
			&nbsp;
			페이지 : <b><%= page %>/ <%= oip.FTotalPage %></b>
		</td>
	</tr>
	<tr align="center" bgcolor="<%= adminColor("tabletop") %>">
   		<td width="20"><input type="checkbox" name="ckall" onclick="ckAll(this)"></td>
		<td width="35">idx<br>(수정)</td>	
		<td width="50">이미지</td>
		<td width="35">상품<br>코드</td>
		<td width="100">브랜드ID</td>
		<td>상품명</td>
		<td width="40">옵션<br>코드</td>
		<td>옵션명</td>	
		<td width="70">상태</td>
		<td width="70">현재재고</td>		
		<td width="35">재고파악시재고</td>
		<td width="35">입력재고</td>
		<td width="35">오차</td>
		<td width="80">재고파악일시</td>
		<td>비고</td>
    </tr>
	<% for i=0 to oip.FResultCount - 1 %>
		<form action="jaegocheck_process.asp" name="frmBuyPrc<%=i%>" method="get">
		<input type="hidden" name="mode">
		<tr bgcolor="#FFFFFF" align="center" onmouseover=this.style.background="orange"; onmouseout=this.style.background='#ffffff';>
		<td><input type="checkbox" name="cksel" onClick="AnCheckClick(this);"></td>
		<td><%= oip.flist(i).fidx %><input type="hidden" name="idx" value="<%= oip.flist(i).fidx %>"></td>	 <!--'인덱스번호 -->
		<td><img src="<%= oip.flist(i).fsmallimage %>" width=50 height=50><input type="hidden" name="smallimage" value="<%= oip.flist(i).fsmallimage %>"></td>	<!--'이미지 -->
		<td><a href="javascript:edit('<%= oip.flist(i).fidx %>','edit')"><%= oip.flist(i).fitemid %></a><input type="hidden" name="itemid" value="<%= oip.flist(i).fitemid %>"></td>				 					<!--'상품번호	 -->
		<td><%= oip.flist(i).fmakerid %><input type="hidden" name="makerid" value="<%= oip.flist(i).fmakerid %>"></td>									 <!--'브랜드id -->
		<td align="left"><%= oip.flist(i).fitemname %><input type="hidden" name="itemname" value="<%= oip.flist(i).fitemname %>"></td>									 <!--'상품명 -->
		<td><%= oip.flist(i).fitemoption %><input type="hidden" name="itemoption" value="<%= oip.flist(i).fitemoption %>"></td>							 <!--'옵션코드 -->
		<td><%= oip.flist(i).fitemoptionname %><input type="hidden" name="itemoptionname" value="<%= oip.flist(i).fitemoptionname %>"></td>				 <!--'옵션명 -->
		<td><%= oip.flist(i).getbigoName %></td>
		<td><%= oip.flist(i).frealstocks %></td>	
		<td><%= oip.flist(i).fbasicstock %><input type="hidden" name="basicstock" value="<%= oip.flist(i).fbasicstock %>"></td>									<!--'재고파악사항 -->
		<td><%= oip.flist(i).frealstock %><input type="hidden" name="realstock" value="<%= oip.flist(i).frealstock %>"></td>							 <!--'현재고파악 -->
										 
		<td><%= oip.flist(i).ferrstock %><input type="hidden" name="errstock" value="<%= oip.flist(i).ferrstock %>"></td>									 <!--'오차	 -->
		<td><%= left(oip.flist(i).factionstartdate,10) %><br><%= mid(oip.flist(i).factionstartdate,11,25) %><input type="hidden" name="actionstartdate" value="<%= oip.flist(i).factionstartdate %>"></td>	 <!--'재고파악일시 -->
			<td>		
		<!--비고란구분시작 -->	
				<% if oip.flist(i).fbigo = 0 then%>		
					<input type="button" value="작업지시" onclick="jak(frmBuyPrc<%=i%>);">-->
				<% elseif oip.flist(i).fbigo =1 then%>
					<input type="button" value="재고입력" class="button" onclick="insert('<%= oip.flist(i).fitemid %>','<%= oip.flist(i).fitemoption %>');">
				<% elseif oip.flist(i).fbigo =5 then%>
					
				<% elseif oip.flist(i).fbigo =7 then%>
					
				<% elseif oip.flist(i).fbigo =8 then%>
					
				<% end if %>
				<input type="hidden" value="<%= oip.flist(i).freguserid %>" name="jisiname"><input type="hidden" value="<%= oip.flist(i).fitemgubun %>" name="itemgubun">
				
			</td>
		<!--비고란구분끝 -->				
		</tr>
		</form>	
	<% next %>
	<% else %>
		<tr bgcolor="#FFFFFF">
			<td colspan="3" align="center" class="page_link">[검색결과가 없습니다.]</td>
		</tr>
	<% end if %>
    <tr height="25" bgcolor="FFFFFF">
		<td colspan="15" align="center">
			<% if oip.HasPreScroll then %>
				<a href="javascript:NextPage('<%= oip.StartScrollPage-1 %>')">[pre]</a>
	   		<% else %>
	    		[pre]
	   		<% end if %>
	
	    	<% for i=0 + oip.StartScrollPage to oip.FScrollCount + oip.StartScrollPage - 1 %>
	    		<% if i>oip.FTotalpage then Exit for %>
		    		<% if CStr(page)=CStr(i) then %>
		    		<font color="red">[<%= i %>]</font>
		    		<% else %>
		    		<a href="javascript:NextPage('<%= i %>')">[<%= i %>]</a>
		    		<% end if %>
	    	<% next %>
	
	    	<% if oip.HasNextScroll then %>
	    		<a href="javascript:NextPage('<%= i %>')">[next]</a>
	    	<% else %>
	    		[next]
    		<% end if %>
		</td>
	</tr>
</table>

<%
set oip = nothing
%>		
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->