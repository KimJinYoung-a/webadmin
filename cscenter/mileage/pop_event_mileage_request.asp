<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description : �̺�Ʈ ���ϸ��� ���� (�����ڿ�)
' Hieditor : 2023.09.18 ������ ����
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
    response.write "<script>alert('�߸��� �����Դϴ�.'); history.back();</script>"
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
            alert("���� ���� ���̵� �Է����ּ���.");
            $("#userid").focus()
        }else if($("#mileage").val()==""){
            alert("���� ���� ���ϸ����� �Է����ּ���.");
            $("#mileage").focus()
        }else if($("#jukyoCD").val()==""){
            alert("���� �ڵ带 �Է����ּ���.");
            $("#jukyoCD").focus()
        }else if($("#jukyo").val()==""){
            alert("���並 �Է����ּ���.");
            $("#jukyo").focus()
        }else{
            $("#mode").val("mileagegive");
            document.frm.submit();
        }
    }

    function fnItemReceive(){
        if($("#userid2").val()==""){
            alert("���� ���� ���̵� �Է����ּ���.");
            $("#userid2").focus()
        }else if($("#itemid").val()==""){
            alert("���� ���� ��ǰ ��ȣ�� �Է����ּ���.");
            $("#itemid").focus()
        }else if($("#eventid").val()==""){
            alert("�̺�Ʈ �ڵ带 �Է����ּ���.");
            $("#eventid").focus()
        }else if($("#give_reason").val()==""){
            alert("���� ������ �Է����ּ���.");
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
                <option value="1">���ϸ���</option>
                <option value="2">��ǰ</option> 
            </select>
        </td>
	</tr>
</table>
<table width="800" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA" id="mileagediv">
	<tr align="left">
  		<td height="30" width="10%" bgcolor="<%= adminColor("tabletop") %>">���̵� :</td>
  		<td bgcolor="#FFFFFF" width="25%" >
  			<b><input type="text" name="userid" id="userid"></b>
  		</td>
  		<td height="30" width="10%" bgcolor="<%= adminColor("tabletop") %>">��������Ʈ :</td>
  		<td bgcolor="#FFFFFF" width="55%"  >
  			<input type="text" name="mileage" id="mileage">
  		</td>
	</tr>
    <tr align="left">
  		<td height="30" width="10%" bgcolor="<%= adminColor("tabletop") %>">���� �ڵ� :</td>
  		<td bgcolor="#FFFFFF" width="25%" >
  			<b><input type="text" name="jukyoCD" id="jukyoCD"></b>
  		</td>
  		<td height="30" width="10%" bgcolor="<%= adminColor("tabletop") %>">���� :</td>
  		<td bgcolor="#FFFFFF" width="55%"  >
  			<input type="text" name="jukyo" id="jukyo" size="55">
  		</td>
	</tr>
    <tr>
  		<td height="30" colspan="4" bgcolor="#FFFFFF" align="center"><input type="button" value="����" onclick="fnMileageReceive()"></td>
	</tr>
</table>
<table width="800" border="0" align="center" class="a" cellpadding="2" cellspacing="1" bgcolor="#BABABA" style="display:none" id="itemdiv">
	<tr align="left">
  		<td height="30" width="10%" bgcolor="<%= adminColor("tabletop") %>">���̵� :</td>
  		<td bgcolor="#FFFFFF" width="25%" >
  			<b><input type="text" name="userid2" id="userid2"></b>
  		</td>
  		<td height="30" width="10%" bgcolor="<%= adminColor("tabletop") %>">��ǰ ��ȣ :</td>
  		<td bgcolor="#FFFFFF" width="55%"  >
  			<input type="text" name="itemid" id="itemid">
  		</td>
	</tr>
    <tr align="left">
  		<td height="30" width="10%" bgcolor="<%= adminColor("tabletop") %>">�ɼ��ڵ� :</td>
  		<td bgcolor="#FFFFFF" width="25%" >
  			<b><input type="text" name="itemoption" id="itemoption" value="0000"></b>
  		</td>
  		<td height="30" width="10%" bgcolor="<%= adminColor("tabletop") %>">���� :</td>
  		<td bgcolor="#FFFFFF" width="55%"  >
  			<input type="text" name="itemea" id="itemea" value="1">
  		</td>
	</tr>
    <tr align="left">
  		<td height="30" width="10%" bgcolor="<%= adminColor("tabletop") %>">�̺�Ʈ ��ȣ :</td>
  		<td bgcolor="#FFFFFF" width="25%" >
  			<b><input type="text" name="eventid" id="eventid"></b>
  		</td>
  		<td height="30" width="10%" bgcolor="<%= adminColor("tabletop") %>">���� ���� :</td>
  		<td bgcolor="#FFFFFF" width="55%"  >
  			<input type="text" name="give_reason" id="give_reason" size="55">
  		</td>
	</tr>
    <tr>
  		<td height="30" colspan="4" bgcolor="#FFFFFF" align="center"><input type="button" value="����" onclick="fnItemReceive()"></td>
	</tr>
</table>
</form>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->