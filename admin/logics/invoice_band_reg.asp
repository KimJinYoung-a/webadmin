<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 송장대역관리
' Hieditor : 2021.04.15 한용민 생성
'###########################################################
%>
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/db_LogisticsOpen.asp" -->
<!-- #include virtual="/admin/incsessionadmin.asp" -->
<!-- #include virtual="/common/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp" -->
<!-- #include virtual="/lib/function.asp" -->
<!-- #include virtual="/lib/offshop_function.asp"-->
<!-- #include virtual="/lib/classes/logistics/invoice_band_cls.asp"-->

<%
dim iidx,siteseq,gubuncd,startsongjangno,endsongjangno,startrealsongjangno,endrealsongjangno
dim remainsongjangcount,basicsongjangyn,isusing,regdate,lastupdate,reguserid,lastuserid, songjangdiv
dim osongjangedit, i, mode, menupos
	menupos = requestcheckvar(getNumeric(request("menupos")),10)
	iidx = requestcheckvar(getNumeric(request("iidx")),10)

mode="add"

set osongjangedit = new cinvoice_band_list
	osongjangedit.frectiidx = iidx

	if iidx <> "" then
		osongjangedit.finvoice_band_one()

        if osongjangedit.FResultCount>0 then
            mode="edit"
            iidx = osongjangedit.FOneItem.fiidx
            siteseq = osongjangedit.FOneItem.fsiteseq
            gubuncd = osongjangedit.FOneItem.fgubuncd
            startsongjangno = osongjangedit.FOneItem.fstartsongjangno
            endsongjangno = osongjangedit.FOneItem.fendsongjangno
            startrealsongjangno = osongjangedit.FOneItem.fstartrealsongjangno
            endrealsongjangno = osongjangedit.FOneItem.fendrealsongjangno
            remainsongjangcount = osongjangedit.FOneItem.fremainsongjangcount
            basicsongjangyn = osongjangedit.FOneItem.fbasicsongjangyn
            isusing = osongjangedit.FOneItem.fisusing
            regdate = osongjangedit.FOneItem.fregdate
            lastupdate = osongjangedit.FOneItem.flastupdate
            reguserid = osongjangedit.FOneItem.freguserid
            lastuserid = osongjangedit.FOneItem.flastuserid
            songjangdiv = osongjangedit.FOneItem.Fsongjangdiv
        end if
	end if
set osongjangedit = nothing

if remainsongjangcount="" or isnull(remainsongjangcount) then remainsongjangcount=0
if basicsongjangyn="" or isnull(basicsongjangyn) then basicsongjangyn="N"
if isusing="" or isnull(isusing) then isusing="Y"
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script type="text/javascript">

	function invoice_band_reg(){
		if ($('#frm select[name="siteseq"] option:selected').val()==''){
			alert('업체를 선택하세요.');
			$('#frm select[name="siteseq"]').focus();
			return;
		}
		if ($('#frm select[name="gubuncd"] option:selected').val()==''){
			alert('출고구분을 선택하세요.');
			$('#frm select[name="gubuncd"]').focus();
			return;
		}
		if ($('#frm input[name="startsongjangno"]').val()==''){
			alert('시작송장번호(검증키포함)를 입력하세요.');
			$('#frm input[name="startsongjangno"]').focus();
			return;
		}
		if ($('#frm input[name="endsongjangno"]').val()==''){
			alert('종료송장번호(검증키포함)를 입력하세요.');
			$('#frm input[name="endsongjangno"]').focus();
			return;
		}
		if ($('#frm input[name="startrealsongjangno"]').val()==''){
			alert('시작실제송장번호를 입력하세요.');
			$('#frm input[name="startrealsongjangno"]').focus();
			return;
		}
		if ($('#frm input[name="endrealsongjangno"]').val()==''){
			alert('종료실제송장번호를 입력하세요.');
			$('#frm input[name="endrealsongjangno"]').focus();
			return;
		}
		if ($('#frm select[name="basicsongjangyn"] option:selected').val()==''){
			alert('기본송장여부를 선택하세요.');
			$('#frm select[name="basicsongjangyn"]').focus();
			return;
		}
		if ($('#frm select[name="isusing"] option:selected').val()==''){
			alert('사용여부를 선택하세요.');
			$('#frm select[name="isusing"]').focus();
			return;
		}
		frm.submit();
	}

</script>

<!-- 액션 시작 -->
<table width="100%" align="center" cellpadding="0" cellspacing="0" class="a" style="padding-top:10;">
<tr>
    <td align="left"></td>
    <td align="right"></td>
</tr>
</table>
<!-- 액션 끝 -->

<form name="frm" id="frm" method="post" action="/admin/logics/invoice_band_process.asp" style="margin:0px;">
<input type="hidden" name="mode" value="<%= mode %>">
<input type="hidden" name="menupos" value="<%= menupos %>">
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="#BABABA">
<% if mode="edit" then %>
<tr bgcolor="#FFFFFF">
    <td align="center">번호</td>
    <td>
        <%= iidx %>
		<input type="hidden" name="iidx" value="<%= iidx %>">
    </td>
</tr>
<% else %>
    <input type="hidden" name="iidx" value="<%= iidx %>">
<% end if %>
<tr bgcolor="#FFFFFF">
    <td align="center">업체</td>
    <td>
        <select class="select" name="siteseq" >
            <option value="10" selected>텐바이텐</option>
        </select>
        <!-- 업체를 추가하려면 온라인출고, 기타출고 모두 수정후에 업체를 추가해야 한다.
        <% drawSelectBoxSiteSeq "siteseq",siteseq,"" %>
        -->
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td align="center">택배사</td>
    <td>
        <% Call drawSelectBoxDeliverCompany ("songjangdiv", songjangdiv) %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td align="center">출고구분</td>
    <td>
        <% drawSelectBoxgubuncd "gubuncd",gubuncd,"" %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td align="center">송장번호(검증키포함)</td>
    <td>
        <input type="text" name="startsongjangno" value="<%= startsongjangno %>" size=11 maxlength=12>
		- <input type="text" name="endsongjangno" value="<%= endsongjangno %>" size=11 maxlength=12>
		<br>맨끝에 1자리는 검증키 입니다. 송장번호와 같이 모두 입력해 주셔야 합니다.
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td align="center">실제송장번호</td>
    <td>
        <input type="text" name="startrealsongjangno" value="<%= startrealsongjangno %>" size=11 maxlength=12>
		- <input type="text" name="endrealsongjangno" value="<%= endrealsongjangno %>" size=11 maxlength=12>
		<br>송장에 노출되는 실제송장번호 입니다. 맨끝에 1자리 검증키는 제외하고 입력해 주세요.
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td align="center">기본송장여부</td>
    <td>
		<% if mode="edit" and basicsongjangyn="Y" then %>
			<input type="hidden" name="basicsongjangyn" value="<%= basicsongjangyn %>">
			<%= basicsongjangyn %>
			<br>기본송장여부를 N으로 수정은 불가합니다.<br>기본송장이 1개 이상 존재 해야 합니다.<br>사용하실 송장대역에서 기본송장여부를 Y 로 해주세요.
		<% else %>
			<% drawSelectBoxisusingYN "basicsongjangyn",basicsongjangyn,"" %>
			<br>Y 일 경우 해당 송장대역으로 로직스에서 출고 됩니다.현재 로직스 실제 사용대역
		<% end if %>
    </td>
</tr>
<tr bgcolor="#FFFFFF">
    <td align="center">사용여부</td>
    <td>
		<% drawSelectBoxisusingYN "isusing",isusing,"" %>
    </td>
</tr>
<% if mode="edit" then %>
	<tr bgcolor="#FFFFFF">
		<td align="center">남은송장수</td>
		<td>
			<%= remainsongjangcount %>
			<br>8시간 주기로 업데이트 됩니다.
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center">최초등록</td>
		<td>
			<%= reguserid %>
			<br><%= regdate %>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="center">최종수정</td>
		<td>
			<%= lastuserid %>
			<br><%= lastupdate %>
		</td>
	</tr>
<% end if %>
<tr bgcolor="#FFFFFF">
	<td align="center" colspan="2">
		<input type="button" value="저장" onclick="invoice_band_reg();" class="button">
	</td>
</tr>
</table>
</form>

<!-- #include virtual="/lib/db/dbclose.asp" -->
<!-- #include virtual="/lib/db/db_LogisticsClose.asp" -->
<!-- #include virtual="/common/lib/poptail.asp"-->
