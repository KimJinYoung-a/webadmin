<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 고객파일전송관리
' History : 2019.11.25 한용민 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/util/md5.asp"-->
<!-- #include virtual="/lib/classes/cscenter/customer_file_cls.asp" -->

<%
Dim menupos , i, page , userhp, userid, orderserial, ccsfileedit, ccsfile, authidx, corderinfo, senduserhp, senduserid, sendorderserial
dim cuserinfo, confirmcertno, filecertsendgubun, asmasteridx, ccsasinfo, sendasmasteridx
	menupos = requestcheckvar(getNumeric(request("menupos")),10)
    authidx = requestcheckvar(getNumeric(request("authidx")),10)
	page = requestcheckvar(getNumeric(request("page")),10)
	userhp = requestcheckvar(request("userhp"),16)
	userid = requestcheckvar(request("userid"),32)
	orderserial = requestcheckvar(request("orderserial"),16)
	filecertsendgubun = requestcheckvar(request("filecertsendgubun"),32)
    asmasteridx = requestcheckvar(getNumeric(request("asmasteridx")),10)

if page = "" then page = 1
if filecertsendgubun="" then filecertsendgubun = "KAKAOTALK"
if userid<>"" and not(isnull(userid)) then
    set cuserinfo = new ccsfilelist
        cuserinfo.frectuserid = userid
        cuserinfo.getuserinfo

        if cuserinfo.ftotalcount>0 then
            senduserid = cuserinfo.FOneItem.fuserid
            senduserhp = cuserinfo.FOneItem.fuserhp
            if userhp="" or isnull(userhp) then userhp = cuserinfo.FOneItem.fuserhp    ' or 문으로 관련 휴대폰번호 전송값 모두 노출
        end if
    set corderinfo = nothing
end if
if orderserial<>"" and not(isnull(orderserial)) then
    set corderinfo = new ccsfilelist
        corderinfo.frectorderserial = orderserial
        corderinfo.getordermasterinfo

        if corderinfo.ftotalcount>0 then
            senduserid = corderinfo.FOneItem.fuserid
            senduserhp = corderinfo.FOneItem.fuserhp
            sendorderserial = corderinfo.FOneItem.forderserial
             if userhp="" or isnull(userhp) then userhp = corderinfo.FOneItem.fuserhp    ' or 문으로 관련 휴대폰번호 전송값 모두 노출
        end if
    set corderinfo = nothing
end if
if asmasteridx<>"" and not(isnull(asmasteridx)) then
    set ccsasinfo = new ccsfilelist
        ccsasinfo.frectasmasteridx = asmasteridx
        ccsasinfo.getcsasinfo

        if ccsasinfo.ftotalcount>0 then
            senduserid = ccsasinfo.FOneItem.fuserid
            senduserhp = ccsasinfo.FOneItem.fuserhp
            sendorderserial = ccsasinfo.FOneItem.forderserial
            sendasmasteridx = ccsasinfo.FOneItem.fasmasteridx
             if userhp="" or isnull(userhp) then userhp = ccsasinfo.FOneItem.fuserhp    ' or 문으로 관련 휴대폰번호 전송값 모두 노출
        end if
    set ccsasinfo = nothing
end if
if userhp<>"" and not(isnull(userhp)) then senduserhp = userhp

set ccsfileedit = new ccsfilelist
	ccsfileedit.frectauthidx = authidx

	if authidx <> "" then
		ccsfileedit.getcsfile_one()
	end if
	
set ccsfile = new ccsfilelist
	ccsfile.FPageSize = 20
	ccsfile.FCurrPage = page
	ccsfile.frectuserhp = trim(userhp)
    ccsfile.frectuserid = trim(userid)
    ccsfile.frectorderserial = trim(orderserial)
    ccsfile.frectasmasteridx = trim(asmasteridx)
	ccsfile.frectisusing = "Y"

    if userhp<>"" or userid<>"" or orderserial<>"" or asmasteridx<>"" then
	    ccsfile.getcsfile()
    end if
%>

<script type="text/javascript">
	function pagesubmit(page){
		frmsearch.page.value = page;
		frmsearch.submit();
	}

	function fileedit(authidx){
		frmedit.authidx.value = authidx;
		frmedit.submit();
	}

	function jumuninput(upfrm){
        if (upfrm.filecertsendgubun.value==''){
            alert('핸드폰번호를 발송할 인증 구분이 없습니다.');
            upfrm.filecertsendgubun.focus();
            return;
        }
        var filecertsendgubun = upfrm.filecertsendgubun.value;

        if (upfrm.senduserhp.value==''){
            alert('인증 받으실 휴대폰번호를 입력 하세요');
            upfrm.senduserhp.focus();
            return;
        }
        var senduserhp = upfrm.senduserhp.value;

        if (confirm('고객님( ' + senduserhp + ' )께 파일 첨부용 링크를 '+ filecertsendgubun + ' 으로 발송 하시겠습니까?')){
            upfrm.mode.value='fileusersend';
            upfrm.action='/cscenter/action/pop_cs_file_send_process.asp';
            upfrm.submit();
        }
	}

</script>

<!-- 검색 시작 -->
<form name="frmsearch" method="get" action="" style="margin:0px;">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="page" value="1">
<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
<tr align="center" bgcolor="#FFFFFF" >
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">검색<br>조건</td>
	<td align="left">
		* 휴대폰번호 : <input type="text" name="userhp" value="<%= userhp %>" size=16 maxlength=16 onKeyPress="if(window.event.keyCode==13) pagesubmit('');">
        &nbsp;&nbsp;
        * 관련고객아이디 : <input type="text" name="userid" value="<%= userid %>" size=16 maxlength=16 onKeyPress="if(window.event.keyCode==13) pagesubmit('');">
        &nbsp;&nbsp;
        * 관련주문번호 : <input type="text" name="orderserial" value="<%= orderserial %>" size=16 maxlength=16 onKeyPress="if(window.event.keyCode==13) pagesubmit('');">
        &nbsp;&nbsp;
        * 관련CS번호 : <input type="text" name="asmasteridx" value="<%= asmasteridx %>" size=16 maxlength=16 onKeyPress="if(window.event.keyCode==13) pagesubmit('');">
	</td>
	<td rowspan="2" width="50" bgcolor="<%= adminColor("gray") %>">
		<input type="button" class="button_s" value="검색" onClick="pagesubmit('');">
	</td>
</tr>
<tr align="center" bgcolor="#FFFFFF" >
	<td align="left">

	</td>
</tr>
</table>
</form>
<!-- 검색 끝 -->
<Br>
<% if userhp="" and userid="" and orderserial="" and asmasteridx="" then %>
    <table width="100%" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="FFFFFF">
		<td align="center" class="page_link"><font color="red">고객정보를 1가지 이상 입력해 주세요.</font></td>
	</tr>
    </table>
<% else %>

    <form name="frmedit" method="post" action="" style="margin:0px;">
    <input type="hidden" name="menupos" value="<%= menupos %>">
    <input type="hidden" name="mode" value="edit">
    <table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <%
    '/수정
    if ccsfileedit.Ftotalcount>0 then
    %>
        <tr bgcolor="#FFFFFF">
            <td align="left" colspan=6>※ 발송내역상세</td>
        </tr>
        <tr bgcolor="#FFFFFF">
            <td align="center" bgcolor="#f1f1f1" width="100">번호</td>
            <td align="left" width="300">
                <%= ccsfileedit.FOneItem.fauthidx %>
                <input type="hidden" size=10 name="authidx" value="<%= ccsfileedit.FOneItem.fauthidx %>">
            </td>
            <td align="center" bgcolor="#f1f1f1" width="100"><font color="red">[필수]</font>휴대폰번호</td>
            <td align="left" width="300">
                <%= ccsfileedit.FOneItem.fuserhp %>
            </td>
            <td align="center" bgcolor="#f1f1f1" width="100">발송기록</td>
            <td align="left">
                <%= ccsfileedit.FOneItem.fregdate %>
                <% if ccsfileedit.FOneItem.fadminid<>"" and not(isnull(ccsfileedit.FOneItem.fadminid)) then %>
                    (<%= ccsfileedit.FOneItem.fadminid %>)
                <% end if %>
                <%
                if C_ADMIN_AUTH then
                    if ccsfileedit.FOneItem.fkakaotalkyn="Y" or ccsfileedit.FOneItem.fsmsyn="Y" then
                %>
                    <%
                    confirmcertno = md5(trim(ccsfileedit.FOneItem.fauthidx) & trim(ccsfileedit.FOneItem.fcertno) & replace(trim(ccsfileedit.FOneItem.fuserhp),"-",""))
                    %>
                    <br>관리자권한 : <% response.write "https://m.10x10.co.kr/cscenter/cs_file_send.asp?nb="& trim(ccsfileedit.FOneItem.fauthidx) &"&certNo="& confirmcertno &"" %>
                <%
                    end if
                end if
                %>
            </td>
        </tr>
        <tr bgcolor="#FFFFFF">
            <td align="center" bgcolor="#f1f1f1">카톡발송</td>
            <td align="left"><%= ccsfileedit.FOneItem.fkakaotalkyn %></td>
            <td align="center" bgcolor="#f1f1f1">문자발송</td>
            <td align="left"><%= ccsfileedit.FOneItem.fsmsyn %></td>
            <td align="center" bgcolor="#f1f1f1">사용여부</td>
            <td align="left"><%= ccsfileedit.FOneItem.fisusing %></td>
        </tr>

        <tr bgcolor="#FFFFFF">
            <td align="center" bgcolor="#f1f1f1">관련고객아이디</td>
            <td align="left"><%= ccsfileedit.FOneItem.fuserid %></td>
            <td align="center" bgcolor="#f1f1f1">관련주문번호</td>
            <td align="left"><%= ccsfileedit.FOneItem.forderserial %></td>
            <td align="center" bgcolor="#f1f1f1">관련CS번호</td>
            <td align="left"><%= ccsfileedit.FOneItem.fasmasteridx %></td>
        </tr>
        <tr bgcolor="#FFFFFF">
            <td align="center" bgcolor="#f1f1f1">상태</td>
            <td align="left"><%= getstatusname(ccsfileedit.FOneItem.fstatus) %></td>
            <td align="center" bgcolor="#f1f1f1">고객파일등록일</td>
            <td align="left" colspan=3><%= ccsfileedit.FOneItem.fcustomer_file_regdate %></td>
        </tr>
        <tr bgcolor="#FFFFFF">
            <td align="center" bgcolor="#f1f1f1">문의내용</td>
            <td align="left" colspan=5>
                <%= nl2br(ccsfileedit.FOneItem.fcomment) %>
            </td>
        </tr>
        <tr bgcolor="#FFFFFF">
            <td align="center" bgcolor="#f1f1f1">첨부파일</td>
            <td align="left" colspan=5>
                <% if trim(ccsfileedit.FOneItem.ffileurl1)<>"" and not(isnull(ccsfileedit.FOneItem.ffileurl1)) then %>
                    <% if instr(ucase(ccsfileedit.FOneItem.ffileurl1),"JPG")>0 or instr(ucase(ccsfileedit.FOneItem.ffileurl1),"JPEG")>0 or instr(ucase(ccsfileedit.FOneItem.ffileurl1),"GIF")>0 or instr(ucase(ccsfileedit.FOneItem.ffileurl1),"PNG")>0 then %>
                        <img src="<%= ccsfileedit.FOneItem.ffileurl1 %>" onfocus="this.blur">
                    <% else %>
                        <a href="<%= ccsfileedit.FOneItem.ffileurl1 %>" target="_blank"><%= GetcsFileName(ccsfileedit.FOneItem.ffileurl1) %>.<%= getFileExtention(ccsfileedit.FOneItem.ffileurl1) %></a>
                    <% end if %>
                <% end if %>
                <% if trim(ccsfileedit.FOneItem.ffileurl2)<>"" and not(isnull(ccsfileedit.FOneItem.ffileurl2)) then %>
                    <Br>
                    <% if instr(ucase(ccsfileedit.FOneItem.ffileurl2),"JPG")>0 or instr(ucase(ccsfileedit.FOneItem.ffileurl2),"JPEG")>0 or instr(ucase(ccsfileedit.FOneItem.ffileurl2),"GIF")>0 or instr(ucase(ccsfileedit.FOneItem.ffileurl2),"PNG")>0 then %>
                        <img src="<%= ccsfileedit.FOneItem.ffileurl2 %>" onfocus="this.blur">
                    <% else %>
                        <a href="<%= ccsfileedit.FOneItem.ffileurl2 %>" target="_blank"><%= GetcsFileName(ccsfileedit.FOneItem.ffileurl2) %>.<%= getFileExtention(ccsfileedit.FOneItem.ffileurl2) %></a>
                    <% end if %>
                <% end if %>
                <% if trim(ccsfileedit.FOneItem.ffileurl3)<>"" and not(isnull(ccsfileedit.FOneItem.ffileurl3)) then %>
                    <Br>
                    <% if instr(ucase(ccsfileedit.FOneItem.ffileurl3),"JPG")>0 or instr(ucase(ccsfileedit.FOneItem.ffileurl3),"JPEG")>0 or instr(ucase(ccsfileedit.FOneItem.ffileurl3),"GIF")>0 or instr(ucase(ccsfileedit.FOneItem.ffileurl3),"PNG")>0 then %>
                        <img src="<%= ccsfileedit.FOneItem.ffileurl3 %>" onfocus="this.blur">
                    <% else %>
                        <a href="<%= ccsfileedit.FOneItem.ffileurl3 %>" target="_blank"><%= GetcsFileName(ccsfileedit.FOneItem.ffileurl3) %>.<%= getFileExtention(ccsfileedit.FOneItem.ffileurl3) %></a>
                    <% end if %>
                <% end if %>
                <% if trim(ccsfileedit.FOneItem.ffileurl4)<>"" and not(isnull(ccsfileedit.FOneItem.ffileurl4)) then %>
                    <Br>
                    <% if instr(ucase(ccsfileedit.FOneItem.ffileurl4),"JPG")>0 or instr(ucase(ccsfileedit.FOneItem.ffileurl4),"JPEG")>0 or instr(ucase(ccsfileedit.FOneItem.ffileurl4),"GIF")>0 or instr(ucase(ccsfileedit.FOneItem.ffileurl4),"PNG")>0 then %>
                        <img src="<%= ccsfileedit.FOneItem.ffileurl4 %>" onfocus="this.blur">
                    <% else %>
                        <a href="<%= ccsfileedit.FOneItem.ffileurl4 %>" target="_blank"><%= GetcsFileName(ccsfileedit.FOneItem.ffileurl4) %>.<%= getFileExtention(ccsfileedit.FOneItem.ffileurl4) %></a>
                    <% end if %>
                <% end if %>
                <% if trim(ccsfileedit.FOneItem.ffileurl5)<>"" and not(isnull(ccsfileedit.FOneItem.ffileurl5)) then %>
                    <Br>
                    <% if instr(ucase(ccsfileedit.FOneItem.ffileurl5),"JPG")>0 or instr(ucase(ccsfileedit.FOneItem.ffileurl5),"JPEG")>0 or instr(ucase(ccsfileedit.FOneItem.ffileurl5),"GIF")>0 or instr(ucase(ccsfileedit.FOneItem.ffileurl5),"PNG")>0 then %>
                        <img src="<%= ccsfileedit.FOneItem.ffileurl5 %>" onfocus="this.blur">
                    <% else %>
                        <a href="<%= ccsfileedit.FOneItem.ffileurl5 %>" target="_blank"><%= GetcsFileName(ccsfileedit.FOneItem.ffileurl5) %>.<%= getFileExtention(ccsfileedit.FOneItem.ffileurl5) %></a>
                    <% end if %>
                <% end if %>
            </td>
        </tr>
    <%
    '/신규발송
    else
    %>
        <input type="hidden" size=10 name="authidx" value="">
        <tr bgcolor="#FFFFFF">
            <td align="left" colspan=6>※ 신규발송</td>
        </tr>
        <tr bgcolor="#FFFFFF">
            <td align="center" bgcolor="#f1f1f1" width=200><font color="red">[필수]</font>휴대폰번호</td>
            <td align="left" colspan=5>
                <input type="text" name="senduserhp"  value="<%= senduserhp %>" size=16 maxlength=16>
                <% drawfilecertsendgubun "filecertsendgubun", filecertsendgubun, "", "N" %>
                <input type="button" value="파일첨부용링크발송" class="button" onclick="jumuninput(frmedit);">
            </td>
        </tr>
        <tr bgcolor="#FFFFFF">
            <td align="center" bgcolor="#f1f1f1" width=200>관련고객아이디</td>
            <td align="left"><input type="text" name="senduserid"  value="<%= senduserid %>" size=16 maxlength=16></td>
            <td align="center" bgcolor="#f1f1f1" width=200>관련주문번호</td>
            <td align="left"><input type="text" name="sendorderserial"  value="<%= sendorderserial %>" size=16 maxlength=16></td>
            <td align="center" bgcolor="#f1f1f1" width=200>관련CS번호</td>
            <td align="left"><input type="text" name="sendasmasteridx"  value="<%= sendasmasteridx %>" size=16 maxlength=16></td>
        </tr>
    <% end if %>

    </table>
    </form>
    <Br>

    <!-- 액션 시작 -->
    <table width="100%" align="center" cellpadding="0" cellspacing="0" class="a">
    <tr>
        <td align="left">
        </td>
        <td align="right">
            <input type="button" class="button" value="신규발송" onClick="pagesubmit('');">
        </td>
    </tr>
    </table>
    <!-- 액션 끝 -->

    <!-- 리스트 시작 -->
    <table width="100%" align="center" cellpadding="0" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
    <tr height="25" bgcolor="FFFFFF">
        <td colspan="20">
            검색결과 : <b><%= ccsfile.FTotalCount %></b>
            &nbsp;
            페이지 : <b><%= page %>/ <%= ccsfile.FTotalPage %></b>
        </td>
    </tr>
    <tr align="center" height="25" bgcolor="<%= adminColor("tabletop") %>">
        <td width=70>번호</td>
        <td width=90>휴대폰번호</td>	
        <td width=30>카톡<br>발송</td>
        <td width=30>문자<br>발송</td>
        <td width=80>발송기록</td>
        <td width=90>관련고객아이디</td>	
        <td width=80>관련주문번호</td>
        <td width=60>상태</td>
        <td>문의내용</td>
        <td width=405>첨부파일</td>
        <td width=40>비고</td>
    </tr>

    <% if ccsfile.FresultCount>0 then %>
        <% for i=0 to ccsfile.FresultCount-1 %>
        <tr align="center" valign="top" bgcolor="<% if cstr(authidx)=cstr(ccsfile.FItemList(i).fauthidx) then %>#f1f1f1<% else %>#FFFFFF<% end if %>">
            <td><%= ccsfile.FItemList(i).fauthidx %></td>
            <td><%= ccsfile.FItemList(i).fuserhp %></td>
            <td><%= ccsfile.FItemList(i).fkakaotalkyn %></td>
            <td><%= ccsfile.FItemList(i).fsmsyn %></td>
            <td>
                <%= left(ccsfile.FItemList(i).fregdate,10) %>
                <br><%= mid(ccsfile.FItemList(i).fregdate,12,16) %>
                <% if ccsfile.FItemList(i).fadminid<>"" and not(isnull(ccsfile.FItemList(i).fadminid)) then %>
                    <br><%= ccsfile.FItemList(i).fadminid %>
                <% end if %>
            </td>
            <td><%= ccsfile.FItemList(i).fuserid %></td>		
            <td><%= ccsfile.FItemList(i).forderserial %></td>
            <td>
                <%= getstatusname(ccsfile.FItemList(i).fstatus) %>
            </td>
            <td align="left"><%= nl2br(ccsfile.FItemList(i).fcomment) %></td>
            <td align="left">
                <% if trim(ccsfile.FItemList(i).ffileurl1)<>"" and not(isnull(ccsfile.FItemList(i).ffileurl1)) then %>
                    <% if instr(ucase(ccsfile.FItemList(i).ffileurl1),"JPG")>0 or instr(ucase(ccsfile.FItemList(i).ffileurl1),"JPEG")>0 or instr(ucase(ccsfile.FItemList(i).ffileurl1),"GIF")>0 or instr(ucase(ccsfile.FItemList(i).ffileurl1),"PNG")>0 then %>
                        <a href="#" onclick="fileedit('<%=ccsfile.FItemList(i).fauthidx %>'); return; false">
                        <img src="<%= ccsfile.FItemList(i).ffileurl1 %>" width=400 height=400 onfocus="this.blur"></a>
                    <% else %>
                        <a href="<%= ccsfile.FItemList(i).ffileurl1 %>" target="_blank"><%= GetcsFileName(ccsfile.FItemList(i).ffileurl1) %>.<%= getFileExtention(ccsfile.FItemList(i).ffileurl1) %></a>
                    <% end if %>
                <% end if %>
                <% if trim(ccsfile.FItemList(i).ffileurl2)<>"" and not(isnull(ccsfile.FItemList(i).ffileurl2)) then %>
                    <br>
                    <% if instr(ucase(ccsfile.FItemList(i).ffileurl2),"JPG")>0 or instr(ucase(ccsfile.FItemList(i).ffileurl2),"JPEG")>0 or instr(ucase(ccsfile.FItemList(i).ffileurl2),"GIF")>0 or instr(ucase(ccsfile.FItemList(i).ffileurl2),"PNG")>0 then %>
                        <a href="#" onclick="fileedit('<%=ccsfile.FItemList(i).fauthidx %>'); return; false">
                        <img src="<%= ccsfile.FItemList(i).ffileurl2 %>" width=400 height=400 onfocus="this.blur"></a>
                    <% else %>
                        <a href="<%= ccsfile.FItemList(i).ffileurl2 %>" target="_blank"><%= GetcsFileName(ccsfile.FItemList(i).ffileurl2) %>.<%= getFileExtention(ccsfile.FItemList(i).ffileurl2) %></a>
                    <% end if %>
                <% end if %>
                <% if trim(ccsfile.FItemList(i).ffileurl3)<>"" and not(isnull(ccsfile.FItemList(i).ffileurl3)) then %>
                    <br>
                    <% if instr(ucase(ccsfile.FItemList(i).ffileurl3),"JPG")>0 or instr(ucase(ccsfile.FItemList(i).ffileurl3),"JPEG")>0 or instr(ucase(ccsfile.FItemList(i).ffileurl3),"GIF")>0 or instr(ucase(ccsfile.FItemList(i).ffileurl3),"PNG")>0 then %>
                        <a href="#" onclick="fileedit('<%=ccsfile.FItemList(i).fauthidx %>'); return; false">
                        <img src="<%= ccsfile.FItemList(i).ffileurl3 %>" width=400 height=400 onfocus="this.blur"></a>
                    <% else %>
                        <a href="<%= ccsfile.FItemList(i).ffileurl3 %>" target="_blank"><%= GetcsFileName(ccsfile.FItemList(i).ffileurl3) %>.<%= getFileExtention(ccsfile.FItemList(i).ffileurl3) %></a>
                    <% end if %>
                <% end if %>
                <% if trim(ccsfile.FItemList(i).ffileurl4)<>"" and not(isnull(ccsfile.FItemList(i).ffileurl4)) then %>
                    <br>
                    <% if instr(ucase(ccsfile.FItemList(i).ffileurl4),"JPG")>0 or instr(ucase(ccsfile.FItemList(i).ffileurl4),"JPEG")>0 or instr(ucase(ccsfile.FItemList(i).ffileurl4),"GIF")>0 or instr(ucase(ccsfile.FItemList(i).ffileurl4),"PNG")>0 then %>
                        <a href="#" onclick="fileedit('<%=ccsfile.FItemList(i).fauthidx %>'); return; false">
                        <img src="<%= ccsfile.FItemList(i).ffileurl4 %>" width=400 height=400 onfocus="this.blur"></a>
                    <% else %>
                        <a href="<%= ccsfile.FItemList(i).ffileurl4 %>" target="_blank"><%= GetcsFileName(ccsfile.FItemList(i).ffileurl4) %>.<%= getFileExtention(ccsfile.FItemList(i).ffileurl4) %></a>
                    <% end if %>
                <% end if %>
                <% if trim(ccsfile.FItemList(i).ffileurl5)<>"" and not(isnull(ccsfile.FItemList(i).ffileurl5)) then %>
                    <br>
                    <% if instr(ucase(ccsfile.FItemList(i).ffileurl5),"JPG")>0 or instr(ucase(ccsfile.FItemList(i).ffileurl5),"JPEG")>0 or instr(ucase(ccsfile.FItemList(i).ffileurl5),"GIF")>0 or instr(ucase(ccsfile.FItemList(i).ffileurl5),"PNG")>0 then %>
                        <a href="#" onclick="fileedit('<%=ccsfile.FItemList(i).fauthidx %>'); return; false">
                        <img src="<%= ccsfile.FItemList(i).ffileurl5 %>" width=400 height=400 onfocus="this.blur"></a>
                    <% else %>
                        <a href="<%= ccsfile.FItemList(i).ffileurl5 %>" target="_blank"><%= GetcsFileName(ccsfile.FItemList(i).ffileurl5) %>.<%= getFileExtention(ccsfile.FItemList(i).ffileurl5) %></a>
                    <% end if %>
                <% end if %>
            </td>
            <td align="center"><input type="button" class="button" value="상세" onclick="fileedit('<%=ccsfile.FItemList(i).fauthidx %>');"></td>
        </tr>
        <% next %>

        <tr height="25" bgcolor="FFFFFF">
            <td colspan="15" align="center">
                <% if ccsfile.HasPreScroll then %>
                    <span class="list_link"><a href="javascript:pagesubmit(<%= ccsfile.StartScrollPage-1 %>);">[pre]</a></span>
                <% else %>
                [pre]
                <% end if %>
                <% for i = 0 + ccsfile.StartScrollPage to ccsfile.StartScrollPage + ccsfile.FScrollCount - 1 %>
                    <% if (i > ccsfile.FTotalpage) then Exit for %>
                    <% if CStr(i) = CStr(ccsfile.FCurrPage) then %>
                    <span class="page_link"><font color="red"><b><%= i %></b></font></span>
                    <% else %>
                    <a href="javascript:pagesubmit(<%= i %>);" class="list_link"><font color="#000000"><%= i %></font></a>
                    <% end if %>
                <% next %>
                <% if ccsfile.HasNextScroll then %>
                    <span class="list_link"><a href="javascript:pagesubmit(<%= i %>);">[next]</a></span>
                <% else %>
                [next]
                <% end if %>
            </td>
        </tr>
    <% else %>
        <tr bgcolor="#FFFFFF">
            <td colspan="20" align="center" class="page_link">[발송된 내역이 없습니다.]</td>
        </tr>
    <% end if %>
    </table>
<% end if %>
<%
set ccsfileedit = nothing
set ccsfile = nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->
