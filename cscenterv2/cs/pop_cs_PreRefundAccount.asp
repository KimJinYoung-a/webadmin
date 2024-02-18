<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/cscenterv2/lib/incSessionAdminCS.asp" -->
<!-- #include virtual="/cscenterv2/lib/db/dbopen.asp" -->
<!-- #include virtual="/cscenterv2/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/cscenterv2/lib/csAsfunction.asp"-->
<!-- #include virtual="/cscenterv2/lib/classes/cs/cs_aslistcls.asp"-->
<%
dim userid, page
userid  = RequestCheckvar(request("userid"),32)
page    = RequestCheckvar(request("page"),10)

if (page="") then page=1

dim frmname, rebankaccount, rebankownername, rebankname
frmname         = RequestCheckvar(request("frmname"),32)
rebankaccount   = RequestCheckvar(request("rebankaccount"),32)
rebankownername = RequestCheckvar(request("rebankownername"),64)
if rebankownername <> "" then
	if checkNotValidHTML(rebankownername) then
	response.write "<script type='text/javascript'>"
	response.write "	alert('유효하지 않은 글자가 포함되어 있습니다. 다시 작성 해주세요');history.back();"
	response.write "</script>"
	response.End
	end if
end if
rebankname      = RequestCheckvar(request("rebankname"),32)

''환불정보
dim orefundInfo
set orefundInfo = New CCSASList
orefundInfo.FCurrpage = page
orefundInfo.FPageSize = 10
orefundInfo.FRectUserID = userid
orefundInfo.GetHisOldRefundInfo


dim i

%>
<body style="margin:10 10 10 10" bgcolor="#FFFFFF">
<script language='javascript'>
var frmname         = "<%= frmname %>";
var rebankname      = "<%= rebankname %>";
var rebankaccount   = "<%= rebankaccount %>";
var rebankownername = "<%= rebankownername %>";


function selectPreAcct(i_rebankname, i_rebankaccount, i_rebankownername){
    eval("opener." + frmname + "." + rebankname).value = i_rebankname;
    eval("opener." + frmname + "." + rebankaccount).value = i_rebankaccount;
    eval("opener." + frmname + "." + rebankownername).value = i_rebankownername;

    window.close();
}

function NextPage(page){
    frm.page.value = page;
    frm.submit();
}
</script>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="0" class="a" bgcolor="#FFFFFF">
<form name="frm" method="get" action="">
<input type="hidden" name="page" value="">
<input type="hidden" name="userid" value="<%= userid %>">
<input type="hidden" name="frmname" value="<%= frmname %>">
<input type="hidden" name="rebankaccount" value="<%= rebankaccount %>">
<input type="hidden" name="rebankaccount" value="<%= rebankaccount %>">
<input type="hidden" name="rebankownername" value="<%= rebankownername %>">
<tr>
    <td>
        <table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
        <tr bgcolor="#FFFFFF">
            <td ><img src="/images/icon_star.gif" align="absbottom">&nbsp; <b>이전 환불 계좌 목록 : <%= userid %></b></td>
        </tr>
        </table>
    </td>
</tr>
</form>
</tr>
    <td>
        <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
        <tr bgcolor="<%= adminColor("topbar") %>">
            <td width="100">은행</td>
            <td width="120">계좌</td>
            <td width="100">예금주</td>
            <td width="60">선택</td>
        </tr>
        <% if orefundInfo.FResultCount>0 then %>
        <% for i=0 to orefundInfo.FResultCount-1 %>
        <tr bgcolor="#FFFFFF">
            <td ><%= orefundInfo.FItemList(i).Frebankname %></td>
            <td ><%= orefundInfo.FItemList(i).Frebankaccount %></td>
            <td ><%= orefundInfo.FItemList(i).Frebankownername %></td>
            <td><input class="button_cs" type="button" value="선택" onClick="selectPreAcct('<%= orefundInfo.FItemList(i).Frebankname %>','<%= orefundInfo.FItemList(i).Frebankaccount %>','<%= orefundInfo.FItemList(i).Frebankownername %>');"></td>
        </tr>
        <% next %>
        <tr bgcolor="#FFFFFF">
            <td colspan="4" align="center">
            <%
				if orefundInfo.HasPreScroll then
					Response.Write "<a href='javascript:NextPage(" & orefundInfo.StarScrollPage-1 & ")'>[pre]</a> &nbsp;"
				else
					Response.Write "[pre] &nbsp;"
				end if

				for i=0 + orefundInfo.StarScrollPage to orefundInfo.FScrollCount + orefundInfo.StarScrollPage - 1

					if i>orefundInfo.FTotalpage then Exit for

					if CStr(page)=CStr(i) then
						Response.Write " <font color='red'>[" & i & "]</font> "
					else
						Response.Write " <a href='javascript:NextPage(" & i & ")'>[" & i & "]</a> "
					end if

				next

				if orefundInfo.HasNextScroll then
					Response.Write "&nbsp; <a href='javascript:NextPage(" & i & ")'>[next]</a>"
				else
					Response.Write "&nbsp; [next]"
				end if
			%>
            </td>
        </tr>
        <% else %>
        <tr bgcolor="#FFFFFF">
            <td colspan="4" align="center">[검색 결과가 없습니다.]</td>
        </tr>
        <% end if %>
        </table>
    </td>
</tr>
</table>
</body>
<%
set orefundInfo = Nothing
%>
<!-- #include virtual="/cscenter/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->