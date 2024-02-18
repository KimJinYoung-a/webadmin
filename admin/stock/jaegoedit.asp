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
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/stockclass/jaegostock.asp"-->

<%
dim fnow,idx, fmode , order , jaego,smallimage,itemid,makerid,itemname,itemoption		'변수선언
dim realstock,basicstock								'변수선언
	idx = html2db(request("idx"))							'테이블의 인덱스값을 받아온다
	fmode = html2db(request("mode")	)						'모드구분
	order = left(now(),10)									'작업지시일
	jaego = html2db(request("jaego"))						'실제재고파악한재고
	smallimage = html2db(request("smallimage"))				'이미지
	itemid = request("itemid")						'상품id
	makerid = html2db(request("makerid"))					'브랜드명
	itemname = html2db(request("itemname"))					'상품명
	itemoption = html2db(request("itemoption"))				'상품옵션코드
	realstock = request("realstock")						'실제재고
	basicstock = request("basicstock")						'재고파악용재고		
	
%>
<% 
dim sql , refer , sql111			'변수선언
%>			

<!--수정모드시작-->
<% if fmode = "edit" then %>
	<%	 
	dim sql101,fitemgubun1,fitemid1,fitemoption1,fitemname1,fitemoptionname1,fmakerid1
	dim fregdate1,freguserid1,forderingdate1,fbasicstock1,frealstock1,ffinishuserid1,fsmallimage1
	
	sql101 = "select"
	sql101 = sql101 & " b.smallimage,b.itemname,b.makerid,b.listimage,"
	sql101 = sql101 & " c.optionname , a.*"
	sql101 = sql101 & " from [db_summary].[dbo].tbl_req_realstock a"
	sql101 = sql101 & " join db_item.[dbo].tbl_item b"
	sql101 = sql101 & " on a.itemid = b.itemid"
	sql101 = sql101 & " left join [db_item].[dbo].tbl_item_option c" 
	sql101 = sql101 & " on a.itemid = c.itemid"
	sql101 = sql101 & " where 1=1 and idx = "& idx &""
	
	'response.write sql101&"<br>"	
	rsget.open sql101,dbget,1
		fitemgubun1 = rsget("itemgubun")				'상품구분
		fitemid1 = rsget("itemid")						'상품번호
		fitemoption1 = rsget("itemoption")		'옵션코드	
		fitemname1 = rsget("itemname")					'상품명
		fitemoptionname1 = rsget("optionname")		'옵션명
		fmakerid1 = rsget("makerid")					'브랜드id
		fregdate1 = rsget("regdate")					'등록일
		freguserid1 = rsget("reguserid")				'지시자id	
		fbasicstock1 = rsget("basicstock")				'재고파악재고
		fsmallimage1 = "http://webimage.10x10.co.kr/image/list/" + GetImageSubFolderByItemid(rsget("itemid")) + "/" + rsget("listimage")				'상품이미지
		frealstock1 = rsget("realstock")				'실사갯수
	rsget.close				
	%>	
	
	<script language="javascript">
	function sendit()
	{
	if(document.form1.jaego.value==""){
	alert("재고파악하신 수량을 입력하세요.")
	document.form1.jaego.focus();
	}
	
	else
	document.form1.mode.value='edit'
	document.form1.submit();
	}
	</script>
	
	<!--표 헤드시작-->
	<body>
	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
		<tr height="10" valign="bottom">
			<td width="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
			<td valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
			<td width="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
		</tr>
		<tr height="25" valign="top">
			<td background="/images/tbl_blue_round_04.gif"></td>
			<td background="/images/tbl_blue_round_06.gif">
				<img src="/images/icon_star.gif" align="absbottom">
				<font color="red" size=2><strong>재고수정 </strong> / 문의사항 : 시스템팀(한용민) </font>
				</td>
				
			<td background="/images/tbl_blue_round_05.gif"></td>
		</tr>
		<tr valign="top">
			<td background="/images/tbl_blue_round_04.gif"></td>
			<td></td>
			<td background="/images/tbl_blue_round_05.gif"></td>
		</tr>
		<tr  height="10" valign="top">
			<td><img src="/images/tbl_blue_round_04.gif" width="10" height="10"></td>
			<td background="/images/tbl_blue_round_06.gif"></td>
			<td><img src="/images/tbl_blue_round_05.gif" width="10" height="10"></td>
		</tr>
		</tr>
	</table>
	<!--표 헤드끝-->
	
	<!--상품테이블시작-->
	<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#CCCCCC">
	 <form method="get" name="form1" action="jaegototalsubmit.asp">  	
	 <input type="hidden" name="itemid" value="<%= fitemid1 %>">
	 <input type="hidden" name="itemoption" value="<%= fitemoption1 %>">
	  <tr bgcolor="#FFFFFF">
	<td rowspan=5><input type="hidden" name="mode"><img src="<%= fsmallimage1 %>" width="100" height="100"></td>
	<td><font size=2>페이지번호 :</font></td>
	 <td><font size=2><%= idx %></font><input type="hidden" name="idx" value="<%= idx %>"></td>
	<td><font size=2>아이템 옵션 : </font></td>
	<td><font size=2><%= fitemoption1 %></font>
	</tr>
	<tr bgcolor="#FFFFFF">
	<td><font size=2>상품번호 :</font></td> 
	<td><font size=2><%= fitemid1 %></font></td>
	<td><font size=2>작업지시자 :</font></td> 
	<td><font size=2><%= freguserid1 %></font></td>
	</tr>
	<tr bgcolor="#FFFFFF">
	<td><font size=2>상품명 : </font></td>
	<td><font size=2><%= fitemname1 %></font></td>
	<td><font size=2>브랜드 :</font></td>
	<td><font size=2><%= fmakerid1 %></font></td>
	</tr>
	<tr bgcolor="#FFFFFF">
	<td><font size=2>상품구분 : </font></td> 
	<td><font size=2>
		<% if fitemgubun1 = 10 then %>
		온라인상품
		<% elseif fitemgubun1 = 90 then %>
			오프라인상품
		<% elseif fitemgubun1 = 70 then %>
			소모품
		<% end if %>
	</font></td>
	<td></td>
	<td></td>
	</tr>
	<tr bgcolor="#FFFFFF">
	<td><font size=2>재고파악시재고 : </font></td>
	<td><font size=2><%= fbasicstock1 %></font><input type="hidden" name="basicstock" value="<%= fbasicstock1 %>"></td>
	<td></td>
	<td></td>
	</tr>
	</table>
	<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#CCCCCC">
	<tr bgcolor="#FFFFFF">
	<td><font size=2>실사재고 : </font> <input type="text" name="jaego" size="12" value="<%= frealstock1 %>"></td>
	<td><input type="button" value="수정" onclick="javascript:sendit()"></tr>
	</tr>
	</form>
	</table>
	<!--상품테이블끝-->
	
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
	</body>
	<!-- 표 하단바 끝-->
	<!--수정모드 끝-->

<!--재고파악모드시작-->
<% elseif fmode = "" then %>					
	
	<%	 
	dim oip1 ,i					'클래스선언
		set oip1 = new Cfitemlist	'변수에 토탈을 넣구
		oip1.Frectitemid = itemid
		if itemoption = "" then
			oip1.frectitemoption = "0000"
		else
			oip1.frectitemoption = itemoption		
		end if
		oip1.fjaegoinsert()			'클래스실행 			
	%>	
	
	<script language="javascript">
	function sendit()
	{
	document.form1.submit();
	}
	</script>
	
	<!--표 헤드시작-->
	<body>
	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a" bgcolor="F4F4F4">
	 <form method="get" name="frm" action=""> 
		<tr height="10" valign="bottom">
			<td width="10" align="right" valign="bottom"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
			<td valign="bottom" background="/images/tbl_blue_round_02.gif"></td>
			<td width="10" align="left" valign="bottom"><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
		</tr>
		<tr height="25" valign="top">
			<td background="/images/tbl_blue_round_04.gif"></td>
			<td background="/images/tbl_blue_round_06.gif">
				<img src="/images/icon_star.gif" align="absbottom">
				<font color="red" size=2><strong>재고파악입력 </strong> </font>
				</td>
				
			<td background="/images/tbl_blue_round_05.gif"></td>
		</tr>
		<tr valign="top">
			<td background="/images/tbl_blue_round_04.gif"></td>
			<td><br>
			상품코드 : <input type="text" name="itemid" value="<%= itemid %>" size="10">
			<a href="javascript:frm.submit();">
			<img src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
			</td>
			<td background="/images/tbl_blue_round_05.gif"></td>
		</tr>
		<tr height="10" valign="top">
			<td><img src="/images/tbl_blue_round_04.gif" width="10" height="10"></td>
			<td background="/images/tbl_blue_round_06.gif"></td>
			<td><img src="/images/tbl_blue_round_05.gif" width="10" height="10"></td>
		</tr>
		</tr>
		</form>
	</table>
	<!--표 헤드끝-->
	
	<% if oip1.ftotalcount > 0 then %>
	<!--상품테이블시작-->
	<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#CCCCCC">
	 <form method="post" name="form1" action="jaegototalsubmit.asp">  	
	  <tr bgcolor="#FFFFFF">
 	  
		<td rowspan=3><img src="<%= oip1.flist(i).fsmallimage %>" width="100" height="100"></td></td>
		<td><font size=2>상품번호 :</font></td>
		<td><font size=2><%= oip1.flist(i).fitemid %><input type="hidden" name="itemid" value="<%= oip1.flist(i).fitemid %>"></font></td>			
		<td><font size=2>아이템 옵션 : </font></td>
		<td><font size=2><%= oip1.flist(i).fitemoption %><input type="hidden" name="itemoption" value="<%= oip1.flist(i).fitemoption %>"></font></td>			
		</tr>
		<tr bgcolor="#FFFFFF">
		<td><font size=2>상품명 : </font></td>
		<td><font size=2><%= oip1.flist(i).fitemname %></font></td>
		<td><font size=2>브랜드 :</font></td>
		<td><font size=2><%= oip1.flist(i).fmakerid %></font></td>
		</tr>
		<tr bgcolor="#FFFFFF">
		<td><font size=2>재고파악용재고 : </font></td>
		<td><font size=2><%= oip1.flist(i).frealstock %></font><input type="hidden" name="basicstock" value="<%= oip1.flist(i).frealstock %>"></td>		<td><font size=2>상품구분 : </font></td> 
		<td><font size=2>
			<% if oip1.flist(i).fitemgubun = 10 then %>
				온라인상품
			<% elseif oip1.flist(i).fitemgubun = 90 then %>
				오프라인상품
			<% elseif oip1.flist(i).fitemgubun = 70 then %>
				소모품
			<% else %>
				<%= oip1.flist(i).fitemgubun %>
			<% end if %>
		</font></td>
		</tr>
	</table>
	<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#CCCCCC">
	<tr bgcolor="#FFFFFF">
	<td><font size=2 colspan=4>재고파악수량 : <input type="text" name="jaego" size="12"> <input type="button" value="저장" onclick="javascript:sendit()"> </font> </tr>
	</tr>
	</form>
	</table>
	<!--상품테이블끝-->
	
	<% else%>
	<table width="100%" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#CCCCCC">
	<tr bgcolor="#FFFFFF">
	<td align=center>[ 검색결과가 없습니다. ]</td></tr>
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
	</body>
	<!-- 표 하단바 끝-->
	<!--재고파악모드끝-->	

<% end if %>

<script language='javascript'>
function GetOnLoad(){
    document.form1.jaego.focus();
    document.form1.jaego.select();
}

window.onload = GetOnLoad;
</script>
<!-- #include virtual="/lib/db/dbclose.asp" -->