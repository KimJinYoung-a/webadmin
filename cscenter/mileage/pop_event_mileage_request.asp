<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : 이벤트 마일리지 적립 (개발자용)
' Hieditor : 2023.09.18 정태훈 생성
'###########################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/lib/db/dbHelper.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/classes/order/new_ordercls.asp"-->
<!-- #include virtual="/lib/classes/cscenter/cs_aslistcls.asp" -->
<!-- #include virtual="/cscenter/lib/csAsfunction.asp"-->
<%
dim UserID
UserID	= session("ssBctID")

if UserID="corpse2" or UserID="tozzinet" or UserID="kobula" then
else
    response.write "<script>alert('잘못된 접속입니다.'); history.back();</script>"
    dbget.close()	:	response.End
end if
%>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script> 
<script>
    function fnReceiveDivSelect(obj){
        if(obj.value==1){
            $("#mileagediv").show();
            $("#itemdiv").hide();
        }else{
            $("#itemdiv").show();
            $("#mileagediv").hide();
        }
    }

    function fnMileageReceive(){
        if($("#userid").val()==""){
            alert("지급 받을 아이디를 입력해주세요.");
            $("#userid").focus()
        }else if($("#mileage").val()==""){
            alert("지급 받을 마일리지를 입력해주세요.");
            $("#mileage").focus()
        }else if($("#jukyoCD").val()==""){
            alert("적요 코드를 입력해주세요.");
            $("#jukyoCD").focus()
        }else if($("#jukyo").val()==""){
            alert("적요를 입력해주세요.");
            $("#jukyo").focus()
        }else{
            $("#mode").val("mileagegive");
            document.frm.submit();
        }
    }

    function fnItemReceive(){
        if($("#userid2").val()==""){
            alert("지급 받을 아이디를 입력해주세요.");
            $("#userid2").focus()
        }else if($("#itemid").val()==""){
            alert("지급 받을 상품 번호를 입력해주세요.");
            $("#itemid").focus()
        }else if($("#eventid").val()==""){
            alert("이벤트 코드를 입력해주세요.");
            $("#eventid").focus()
        }else if($("#give_reason").val()==""){
            alert("지급 사유를 입력해주세요.");
            $("#give_reason").focus()
        }else{
            $("#mode").val("itemgive");
            document.frm.submit();
        }
    }
</script>
<form name="frm" method="post" action="/cscenter/mileage/doEventGive.asp" onsubmit="return false;" style="margin:0px;">
<input type="hidden" name="mode" id="mode" value="request">
<table width="800" border="0" align="center" class="a" cellpadding="0" cellspacing="0">
    <tr align="left">
  		<td height="30">
            <select name="receiveDiv" onchange="fnReceiveDivSelect(this)">
                <option value="1">마일리지</option>
                <option value="2">상품</option> 
            </select>
        </td>
	</tr>
</table>
<table width="800" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA" id="mileagediv">
	<tr align="left">
  		<td height="30" width="10%" bgcolor="<%= adminColor("tabletop") %>">아이디 :</td>
  		<td bgcolor="#FFFFFF" width="25%" >
  			<b><input type="text" name="userid" id="userid"></b>
  		</td>
  		<td height="30" width="10%" bgcolor="<%= adminColor("tabletop") %>">적립포인트 :</td>
  		<td bgcolor="#FFFFFF" width="55%"  >
  			<input type="text" name="mileage" id="mileage">
  		</td>
	</tr>
    <tr align="left">
  		<td height="30" width="10%" bgcolor="<%= adminColor("tabletop") %>">적요 코드 :</td>
  		<td bgcolor="#FFFFFF" width="25%" >
  			<b><input type="text" name="jukyoCD" id="jukyoCD"></b>
  		</td>
  		<td height="30" width="10%" bgcolor="<%= adminColor("tabletop") %>">적요 :</td>
  		<td bgcolor="#FFFFFF" width="55%"  >
  			<input type="text" name="jukyo" id="jukyo" size="55">
  		</td>
	</tr>
    <tr>
  		<td height="30" colspan="4" bgcolor="#FFFFFF" align="center"><input type="button" value="지급" onclick="fnMileageReceive()"></td>
	</tr>
</table>
<table width="800" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA" style="display:none" id="itemdiv">
	<tr align="left">
  		<td height="30" width="10%" bgcolor="<%= adminColor("tabletop") %>">아이디 :</td>
  		<td bgcolor="#FFFFFF" width="25%" >
  			<b><input type="text" name="userid2" id="userid2"></b>
  		</td>
  		<td height="30" width="10%" bgcolor="<%= adminColor("tabletop") %>">상품 번호 :</td>
  		<td bgcolor="#FFFFFF" width="55%"  >
  			<input type="text" name="itemid" id="itemid">
  		</td>
	</tr>
    <tr align="left">
  		<td height="30" width="10%" bgcolor="<%= adminColor("tabletop") %>">옵션코드 :</td>
  		<td bgcolor="#FFFFFF" width="25%" >
  			<b><input type="text" name="itemoption" id="itemoption" value="0000"></b>
  		</td>
  		<td height="30" width="10%" bgcolor="<%= adminColor("tabletop") %>">수량 :</td>
  		<td bgcolor="#FFFFFF" width="55%"  >
  			<input type="text" name="itemea" id="itemea" value="1">
  		</td>
	</tr>
    <tr align="left">
  		<td height="30" width="10%" bgcolor="<%= adminColor("tabletop") %>">이벤트 번호 :</td>
  		<td bgcolor="#FFFFFF" width="25%" >
  			<b><input type="text" name="eventid" id="eventid"></b>
  		</td>
  		<td height="30" width="10%" bgcolor="<%= adminColor("tabletop") %>">지급 사유 :</td>
  		<td bgcolor="#FFFFFF" width="55%"  >
  			<input type="text" name="give_reason" id="give_reason" size="55">
  		</td>
	</tr>
    <tr>
  		<td height="30" colspan="4" bgcolor="#FFFFFF" align="center"><input type="button" value="지급" onclick="fnItemReceive()"></td>
	</tr>
</table>
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->