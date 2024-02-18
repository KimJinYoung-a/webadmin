<%@ Language=VBScript %>
<%option explicit%>
<!-- #include virtual="/common/incSessionBctId.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<HTML>
<HEAD>
<TITLE>우편번호 검색 </TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link rel=stylesheet type="text/css" href="/css/scm.css">
<style type="text/css">
.input {
	font-family: "돋움", "Verdana";
	font-size: 12px;
	line-height: 20px;
	color: #000000;
	background-color: #FFFFFF;
	border: 0px solid #FFFFFF;
}
.kindTab_on {width:50%;background-color:#DEF8D0;padding:5px;float:left;font-weight:bold;cursor:pointer;}
.kindTab_off {width:50%;background-color:#E0E0E0;padding:5px;float:left;font-weight:normal;cursor:pointer;}
.selBtn {font-size: 12px;border:1px;}
</style>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
</HEAD>
<body bgcolor="white" text="black" link="black" vlink="black" alink="#464646" style="margin:0 0 0 0">
<script type="text/javascript">
<!--
	function SubmitForm(frm) {
		if (frm.query.value.length < 2) { alert("동 이름을 두글자 이상 입력하세요."); return; }
		frm.submit();
	}

	function Copy(t,post1,post2,add,dong) {
		opener.CopyZip(t,post1,post2,add,dong)
		window.close();
	}

	function chgTab(dv) {
		if(dv=="a") {
			$("#kTab1").attr("class","kindTab_on");
			$("#kTab2").attr("class","kindTab_off");
			$("#dRow1").show();
			$("#dRow2").hide();
			$("#stype").val("addr");
		} else {
			$("#kTab1").attr("class","kindTab_off");
			$("#kTab2").attr("class","kindTab_on");
			$("#dRow1").hide();
			$("#dRow2").show();
			$("#stype").val("road");
		}
	}
//-->
</script>
<%
	' -------------------------------------
	' 회원의 주소를 찾는 Popup Window 화면
	' -------------------------------------
	Dim strTarget, stype
	Dim strQuery

	strTarget	= requestCheckVar(Request("target"),32)
	strQuery	= requestCheckVar(Request("query"),20)
	stype		= requestCheckVar(Request("stype"),20)

	If strQuery = "" then
%>
	<table width="440" border="0" cellpadding="0" cellspacing="0" class="a">
	<tr>
		<td background="/fiximage/web2007/zipcode/member_pop_bg.gif" height="55">
			<div align="center"><img src="/fiximage/web2007/zipcode/search_zip.gif" width="114" height="36"></div>
		</td>
	</tr>
	<tr>
		<td>
			<div align="center" id="kTab1" class="kindTab_on" onclick="chgTab('a')">지번검색</div>
			<div align="center" id="kTab2" class="kindTab_off" onclick="chgTab('r')">도로명검색</div>
		</td>
	</tr>
	<tr id="dRow1">
		<td height="50">
			<div align="center">찾고자 하는 주소의 동/읍/면 이름을 입력하세요.<br>
			(예: 대치동,곡성음,오곡면)</div>
		</td>
	</tr>
	<tr id="dRow2" style="display:none;">
		<td height="50">
			<div align="center">찾고자 하는 주소의 도로명 이름을 입력하세요.<br>
			(예: 동숭1길, 세종대로)</div>
		</td>
	</tr>
	<tr>
		<td height="37">
			<div align="center">
			<form action="/lib/searchzip3.asp?" method="get" name="gil" onsubmit="SubmitForm(document.gil); return false;" style="margin:0px;">
			<input type="hidden" name="mode"	value="search">
			<input type="hidden" name="target"	value="<%=strTarget%>">
			<input type="hidden" name="form"	value="account">
			<input type="hidden" name="post1"	value="post1">
			<input type="hidden" name="post2"	value="post2">
			<input type="hidden" name="add"		value="add">
			<input type="hidden" name="stype"	id="stype" value="addr">
			<table border="0" cellpadding="0" class="a">
			<tr>
				<td>검색어 :</td>
				<td width="97">
					<input type="text" name="query" class="input_01" size="15">
				</td>
				<td width="61"><a href="javascript:SubmitForm(document.gil);"><img src="/fiximage/web2007/zipcode/zip_search.gif" width="65" height="22" border="0"></a></td>
			</tr>
			</table>
			</form>
			</div>
		</td>
	</tr>
	<tr>
		<td height="8">
			<div align="center"><img src="/fiximage/web2007/zipcode/pup_dotline.gif" width="144" height="8"></div>
		</td>
	</tr>
	<tr>
		<td height="37">
			<div align="center">
			<table width="380" border="0" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC">
			<tr>
				<td>
					<table width="380" border="0" cellpadding="2" cellspacing="1">
					<tr bgcolor="#f7f7f7">
						<td class="a" width="109">
							<div align="center"><b><font color="#666666">우편번호</font></b></div>
						</td>
						<td class="a" width="290">
							<div align="center"><b><font color="#666666">주소</font></b></div>
						</td>
						<td class="a" width="50" bgcolor="#f7f7f7">
							<div align="center"><b><font color="#666666">선택</font></b></div>
						</td>
					</tr>
					<tr bgcolor="#FFFFFF">
						<td colspan="3" class="a">
							<div align="center"><font color="#999999">검색어를 입력해주세요</font></div>
						</td>
					</tr>
					</table>
				</td>
			</tr>
			</table>
			</div>
		</td>
	</tr>
	<tr>
		<td style="border-bottom: 7px solid #e1e1e1;" height="38">
			<div align="center"><a href="javascript:self.close();"><img src="/fiximage/web2007/zipcode/zip_close.gif" width="65" height="22" border="0"></a></div>
		</td>
	</tr>
	</table>
<%	else %>
	<table width="440" border="0" cellpadding="0" cellspacing="0" class="a">
	<tr>
		<td background="/fiximage/web2007/zipcode/member_pop_bg.gif" height="55">
			<div align="center"><img src="/fiximage/web2007/zipcode/search_zip.gif" width="114" height="36"></div>
		</td>
	</tr>
	<tr>
		<td>
			<div align="center" id="kTab1" class="<%=chkIIF(stype="addr","kindTab_on","kindTab_off")%>" onclick="chgTab('a')">지번검색</div>
			<div align="center" id="kTab2" class="<%=chkIIF(stype="road","kindTab_on","kindTab_off")%>" onclick="chgTab('r')">도로명검색</div>
		</td>
	</tr>
	<tr id="dRow1" style="<%=chkIIF(stype="addr","","display:none")%>">
		<td height="50">
			<div align="center">찾고자 하는 주소의 동/읍/면 이름을 입력하세요.<br>
			(예: 대치동,곡성음,오곡면)</div>
		</td>
	</tr>
	<tr id="dRow2" style="<%=chkIIF(stype="road","","display:none")%>">
		<td height="50">
			<div align="center">찾고자 하는 주소의 도로명 이름을 입력하세요.<br>
			(예: 동숭1길, 세종대로)</div>
		</td>
	</tr>
	<tr>
		<td  height="37">
			<div align="center">
			<form action="/lib/searchzip3.asp?" method="get" name="gil2" onsubmit="SubmitForm(document.gil2); return false;" style="margin:0px;">
			<input type="hidden" name="target" value="<%=strTarget%>">
			<input type="hidden" name="stype"	id="stype" value="<%=stype%>">
				<table border="0" cellpadding="0" class="a">
				<tr>
					<td>검색어 :</td>
					<td width="97">
						<input type="text" name="query" class="input_01" size="15">
					</td>
					<td width="61"><a href="javascript:SubmitForm(document.gil2);"><img src="/fiximage/web2007/zipcode/zip_search.gif" width="65" height="22" border="0"></a></td>
				</tr>
				</table>
			</form>
			</div>
		</td>
	</tr>
	<tr>
		<td height="8">
			<div align="center"><img src="/fiximage/web2007/zipcode/pup_dotline.gif" width="144" height="8"></div>
		</td>
	</tr>
	<tr>
		<td height="37">
			<div align="center">
			<table width="400" border="0" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC" class="a">
			<tr>
				<td>
					<table width="400" border="0" cellpadding="2" cellspacing="1">
					<tr>
						<td class="a" width="109" bgcolor="#f7f7f7">
							<div align="center"><b><font color="#666666">우편번호</font></b></div>
						</td>
						<td class="a" width="290" bgcolor="#f7f7f7">
							<div align="center"><b><font color="#666666">주소</font></b></div>
						</td>
						<td class="a" width="50" bgcolor="#f7f7f7">
							<div align="center"><b><font color="#666666">선택</font></b></div>
						</td>
					</tr>
		<%
			Dim strSql
			Dim nRowCount

			Dim strAddress

			dim useraddr01, useraddr02

			dim lstr
		        lstr = CStr(Len(strQuery))

				''if stype="addr" then
				''	strSql = "SELECT   ADDR_ZIP1, ADDR_ZIP2, ADDR_SI,ADDR_GU,ADDR_DONG,ADDR_ETC,ADDR_Fulltext FROM [db_zipcode].[dbo].ADDR080TL WHERE ADDR_Fulltext like '%" & strQuery & "%' and ADDR_sortNo<>'999' "
				''elseif stype="road" then
				''	strSql = "SELECT   ADDR_ZIP1, ADDR_ZIP2, ADDR_SI,ADDR_GU,ADDR_DONG,ADDR_ROAD,ADDR_BLDNO1,ADDR_BLDNO2,ADDR_ETC,ADDR_Fulltext " &_
				''			" FROM [db_zipcode].[dbo].ROAD010 " &_
				''			" WHERE ADDR_ROAD like '" & strQuery & "%' and ADDR_sortNo<>'999' " &_
				''			" order by addr_zip1, addr_Gu, addr_Road, Addr_BldNo1 "
				''end if

				strSql = " [db_zipcode].[dbo].[usp_Ten_GetZipcodeList] '" + CStr(strQuery) + "', '" + CStr(stype) + "' "

				rsget.Open strSQL,dbget,1
				'oRs.Open strSQL,oCnn,1

				if not rsget.eof then
					do while not rsget.EOF

						if stype="addr" then
							strAddress = trim(rsget("ADDR_Fulltext"))

							useraddr01 = trim(rsget("ADDR_SI")) & " " & trim( rsget("ADDR_GU"))
							useraddr02 = trim( rsget("ADDR_DONG")) & " " & trim( rsget("ADDR_ETC"))
							useraddr02 = Replace(useraddr02,"'","\'")
						elseif stype="road" then
							strAddress = trim(rsget("ADDR_Fulltext")) & " " & trim( rsget("ADDR_BLDNO1"))
							if Not(rsget("ADDR_BLDNO2")="" or isNull(rsget("ADDR_BLDNO2"))) then
								strAddress = strAddress & " ~ " & trim(rsget("ADDR_BLDNO2"))
							end if

							useraddr01 = trim(rsget("ADDR_SI")) & " " & trim( rsget("ADDR_GU"))
							
							'' 읍면동 추가 (택배사 주소정제 관련(2016/07/07)
							''useraddr02 = trim( rsget("ADDR_ROAD"))
							useraddr02 = trim( rsget("ADDR_DONG")) & " " & trim( rsget("ADDR_ROAD"))
							if Not(rsget("ADDR_ETC")="" or isNull(rsget("ADDR_ETC"))) then
								'다량 배송처가 있는 곳은 단일 건물
								useraddr02 = useraddr02 & " " & trim(rsget("ADDR_BLDNO1")) & " " & trim(rsget("ADDR_ETC"))
							end if
							useraddr02 = Replace(useraddr02,"'","\'")
						end if
		%>
					<tr bgcolor="#FFFFFF">
						<td class="a" width="109" align="center" onclick="Copy('<%= strTarget %>','<%=rsget("ADDR_ZIP1")%>','<%=rsget("ADDR_ZIP2")%>','<% = useraddr01 %>', '<% = useraddr02 %>')" style="cursor:hand">
							<input type="text" name="post1" size="3" value='<%=rsget("ADDR_zip1")%>' class="input" style="cursor:hand"> -
							<input type="text" name="post2" size="3" value='<%=rsget("ADDR_zip2")%>' class="input" style="cursor:hand">
						</td>
						<td class="a" width="290"  align="center">
							<INPUT type="text" name="add" value='<%=strAddress%>' size="42" class="input">
						</td>
						<td class="a" width="50" >
							<div align="center"><input type="button" class="selBtn" onclick="Copy('<%= strTarget %>','<%=rsget("ADDR_ZIP1")%>','<%=rsget("ADDR_ZIP2")%>','<% = useraddr01 %>', '<% = useraddr02 %>')" value="선택"></div>
						</td>
					</tr>
		<%
						rsget.MoveNext
					loop
				end if
		%>
					</table>
				</td>
			</tr>
			</table>
			</div>
		</td>
	</tr>
	<tr>
		<td style="border-bottom: 7px solid #e1e1e1;" height="38">
			<div align="center"><a href="javascript:self.close();"><img src="/fiximage/web2007/zipcode/zip_close.gif" width="65" height="22" border="0"></a></div>
		</td>
	</tr>
	</table>
<%	end if %>
<!-- #include virtual="/lib/db/dbclose.asp" -->
