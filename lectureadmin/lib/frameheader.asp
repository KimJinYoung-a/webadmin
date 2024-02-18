<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/classes/partners/fingersUpcheAgreeCls.asp"-->
<%
dim btcid,grpid
btcid= session("ssBctID")
grpid= session("ssGroupid")
if (btcid="") then response.End

''계약/약관 동의 관련 체크
dim isAgreeReq : isAgreeReq = false
if (session("isAgreeReq")="") then
    isAgreeReq = IsFingersUpcheAgreeNotiRequire(grpid,btcid)
elseif (session("isAgreeReq")="Y") then
    isAgreeReq = true
end if
%>

<html>
<head>
<title>[10x10] Business Comunication</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link rel="stylesheet" href="/css/scm.css" type="text/css">
<script language='javascript'>
function getonload(){
    <% if (isAgreeReq) then %>
    //alert('판매자 이용 약관 및 계약서 동의 후 이용 가능합니다.');
    top.contents.location.href="/lectureadmin/contract/ctrListBrand.asp?menupos=1816";
    <% End if %>
}

function WindowMinSize(){
	parent.document.all('menuset').cols = "20,*";
	document.all.WINSIZE[0].style.display = "none";
	document.all.WINSIZE[1].style.display = "";
}

function WindowMaxSize(){
	parent.document.all('menuset').cols = "180,*";
	document.all.WINSIZE[0].style.display = "";
	document.all.WINSIZE[1].style.display = "none";
}

function pop_editcompany(){
	var popwin = window.open('/designer/company/editcompany3.asp?menupos=53' ,'op1','width=750,height=600,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function pop_10x10_person(){
	var popwin = window.open('/common/pop_10x10_person.asp','op2','width=450,height=450,scrollbars=yes,resizable=yes');
	popwin.focus();
}

function pop_10x10_map(){
	var popwin = window.open('/common/pop_10x10_map.asp','op3','width=650,height=800,scrollbars=yes,resizable=yes')
	popwin.focus();
}

function ShiftBrand(comp){
    top.contents.location.href="/designer/lib/shiftbrand.asp?shiftid="+comp.value;
    <% if (FALSE) then ''2016/08/11  %>
	// refere 때문에 동적으로 생성
	var targetFrm = top.contents;


	var o  = targetFrm.document.createElement("form");
    var oi1 = targetFrm.document.createElement("input");

	oi1.type = "hidden";
    oi1.name = "shiftid";
    oi1.value = comp.value;

    o.appendChild(oi1);
    targetFrm.document.body.appendChild(o);

    o.method = "get";
    o.action = "/designer/lib/shiftbrand.asp";



    o.submit();


	//focusing out //주석처리 2016/08/08
	//document.location.reload();
    <% end if %>
    
}
</script>



</head>

<!-- 상단 여백 -->
<body bgcolor="#FFFFFF" text="#000000" topmargin="0" leftmargin="0" onload="getonload()">

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
	<tr height="5">
		<td></td>
	</tr>
</table>
<!-- 상단 여백 -->


<!-- 표 상단바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
   	<tr height="10" valign="bottom" bgcolor="F4F4F4">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="top" bgcolor="F4F4F4">
	        <td background="/images/tbl_blue_round_04.gif"></td>
	        <td valign="center" bgcolor="F4F4F4">
	        	<img src="/images/admin_logo_10x10.jpg" width="90" height="25" align="absbottom">
	        	<b>10x10 Business Communication Tool</b>
	        </td>
	        <td valign="center" align="right" bgcolor="F4F4F4">
	        <%
dim sqlStr,i
dim Resultcount


sqlStr = " select top 30 p.id, c.socname, c.socname_kor,c.userdiv" + VbCrlf
sqlStr = sqlStr + " from [db_partner].[dbo].tbl_partner p" + VbCrlf
sqlStr = sqlStr + " left join [db_user].[dbo].tbl_user_c c on p.id=c.userid" + VbCrlf
sqlStr = sqlStr + " where p.groupid='" + grpid + "'" + VbCrlf
sqlStr = sqlStr + " and p.userdiv='9999'" + VbCrlf
sqlStr = sqlStr + " and p.isusing='Y'" + VbCrlf
sqlStr = sqlStr + " and c.userdiv<15"  ''매입처만.. => 강사포함 (14)

''sqlStr = sqlStr + " and c.isusing='Y'" + VbCrlf

rsget.Open sqlStr,dbget,1

	if not rsget.Eof then
		Resultcount = rsget.RecordCount
%>
        	<select class="select" name="brandshift" onChange="ShiftBrand(this)">
        	<% for i=0 to Resultcount - 1 %>
        	<option value="<%= rsget("id") %>" <% if (LCase(rsget("id"))=LCase(session("ssBctId"))) then response.write "selected" %> ><%= rsget("id") %> (<%= db2html(rsget("socname_kor")) %> <%= CHKIIF(rsget("userdiv")="14","-더핑거스","") %>)
        	<% rsget.MoveNext %>
        	<% next %>
        	</select>
        	&nbsp;
<%
	end if
rsget.Close
%>
	        
	        	<a href="javascript:pop_editcompany('<%= menupos %>');" onMouseOver="this.style.color = 'red'; this.style.fontWeight = 'bold'" onMouseOut="this.style.color = 'black'; this.style.fontWeight = 'normal'">업체 및 브랜드정보수정</a>
		        <img src="/images/barDupe_18px.gif" width="2" height="18" align="absbottom">
		        <a href="javascript:pop_10x10_person();" onMouseOver="this.style.color = 'red'; this.style.fontWeight = 'bold'" onMouseOut="this.style.color = 'black'; this.style.fontWeight = 'normal'">파트별 담당자</a>
		        <img src="/images/barDupe_18px.gif" width="2" height="18" align="absbottom">
		        <a href="javascript:pop_10x10_map();" onMouseOver="this.style.color = 'red'; this.style.fontWeight = 'bold'" onMouseOut="this.style.color = 'black'; this.style.fontWeight = 'normal'">텐바이텐 약도</a>
            </td>
	        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
</table>
<!-- 표 상단바 끝-->

<!-- 표 중간바 시작-->
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
    <tr bgcolor="#CCCCCC" height="20">
        <td width="10" align="right" background="/images/tbl_blue_round_04.gif"></td>
        <td width="170" align="right">
			<div id=WINSIZE style="display:">창 확대하기
				<input type=button value="☜" onClick="javascript:WindowMinSize()">
			</div>
			<div id=WINSIZE style="display:none">창 축소하기
				<input type=button value="☞" onClick="javascript:WindowMaxSize()">
			</div>
		</td>
        <td align="right">
	        <b><%=session("ssBctID")%>(<%=session("ssBctCname")%>)</b> 님이 로그인 하셨습니다.
	    	&nbsp;
	    	<a href="/login/dologout.asp" target="_top"><img src="/images/icon_logout.gif" width="64" height="17" border="0" align="absbottom"></a>
        </td>
        <td width="10" align="left" background="/images/tbl_blue_round_05.gif"></td>
    </tr>
</table>
<!-- 표 중간바 끝-->



</body>
</html>
<!-- #include virtual="/lib/db/dbclose.asp" -->