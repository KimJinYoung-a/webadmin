<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/items/newitemcouponcls.asp"-->
<!-- #include virtual="/lib/function.asp" -->
<%
dim itemcouponidx
dim oitemcouponmaster
dim IsEditMode, IsExpiredCoupon

itemcouponidx = requestCheckVar(request("itemcouponidx"),9)
if itemcouponidx="" then itemcouponidx=0

set oitemcouponmaster = new CItemCouponMaster
oitemcouponmaster.FRectItemCouponIdx = itemcouponidx
oitemcouponmaster.GetOneItemCouponMaster

IsEditMode = (CStr(itemcouponidx)<>"0")
%>

<script type="text/javascript" src="https://cdn.jsdelivr.net/npm/es6-promise@4/dist/es6-promise.min.js"></script>
<script type="text/javascript" src="https://cdn.jsdelivr.net/npm/es6-promise@4/dist/es6-promise.auto.min.js"></script>
<script type="text/javascript" src="/js/jquery-1.7.1.min.js"></script>
<script language='javascript'>
    let changeCouponImageFlag = false;

    function OpenCouponMaster(){
        frmcoupon.mode.value="opencoupon";

        if (confirm('������ ���� �Ͻðڽ��ϱ�?')){
            frmcoupon.submit();
        }
    }

    function reserveCouponMaster(){
        frmcoupon.mode.value="reservecoupon";

        if (confirm('���������� ���� �Ͻðڽ��ϱ�?')){
            frmcoupon.submit();
        }

    }

    var alertCnt = 0;
    function AlertMarginChange(){
        if (alertCnt==0){
            alert('���� ������ �����Ͻø� ����ǰ ��ü�� ���� �˴ϴ�.');
            alertCnt++;
        }
    }

    function CloseCouponMaster(){
        frmcoupon.mode.value="closecoupon";

        if (confirm('!! ���� ����� ������ ���� ���� �˴ϴ�.\n\n������ ���� ���� �Ͻðڽ��ϱ�?')){
            frmcoupon.submit();
        }
    }
    function fninput(v){

        var ele = document.getElementById('marginlayer');

        if (v==20){
            ele.style.display="";
        }else {
            ele.style.display="none";
        }
    }

    function SaveCouponMaster(frm, isEditMode){
        if (frmcoupon.itemcouponname.value.length<2){
            alert('�������� �Է��� �ּ���.');
            frmcoupon.itemcouponname.focus();
            return;
        }

        if ((!frmcoupon.couponGubun[0].checked)&&(!frmcoupon.couponGubun[1].checked)&&(!frmcoupon.couponGubun[2].checked)&&(!frmcoupon.couponGubun[3].checked)){
            alert('���� ������ �����ϼ���..');
            frmcoupon.couponGubun[0].focus();
            return;
        }

        if (frmcoupon.couponGubun[2].checked){
            alert('���� ���� ����� �ý�����  ���� ���!');
        }

        if (frmcoupon.itemcouponvalue.value.length<1){
            alert('���� �ݾ� �Ǵ� �������� �Է��� �ּ���.');
            frmcoupon.itemcouponvalue.focus();
            return;
        }

        if (!IsDigit(frmcoupon.itemcouponvalue.value)){
            alert('���� �ݾ� �Ǵ� �������� ���ڸ� �����մϴ�.');
            frmcoupon.itemcouponvalue.focus();
            return;
        }


        if ((!frmcoupon.itemcoupontype[0].checked)&&(!frmcoupon.itemcoupontype[1].checked)&&(!frmcoupon.itemcoupontype[2].checked)){
            alert('���� Ÿ���� ������ �ּ���.');
            frmcoupon.itemcouponvalue.focus();
            return;
        }

        if ((frmcoupon.itemcoupontype[2].checked)&&(frmcoupon.itemcouponvalue.value!='<%=Cstr(getDefaultBeasongPayByDate(now()))%>')){
            alert('������ ������ ���ξ��� <%=Cstr(getDefaultBeasongPayByDate(now()))%>�� �Դϴ�.');
            frmcoupon.itemcouponvalue.focus();
            return;
        }

        if ((frmcoupon.itemcoupontype[2].checked)&&!(frmcoupon.margintype.value=='20'||frmcoupon.margintype.value=='50'||frmcoupon.margintype.value=='80')){
    //		alert('������ ���� �߱޽� �ݹݺδ�, �������� �Ǵ� ������500��ü�δ����� �������ּ���.');
    //		frmcoupon.margintype.focus();
    //		return;
        }


        if (frmcoupon.itemcouponstartdate.value.length!=10){
            alert('���� �߱� �������� �Է��� �ּ���.');
            frmcoupon.itemcouponstartdate.focus();
            return;
        }

        if (frmcoupon.itemcouponstartdate2.value.length!=8){
            alert('���� �߱� �������� �Է��� �ּ���.');
            frmcoupon.itemcouponstartdate2.focus();
            return;
        }

        if (frmcoupon.itemcouponexpiredate.value.length!=10){
            alert('���� �߱� �������� �Է��� �ּ���.');
            frmcoupon.itemcouponexpiredate.focus();
            return;
        }

        if (frmcoupon.itemcouponexpiredate2.value.length!=8){
            alert('���� �߱� �������� �Է��� �ּ���.');
            frmcoupon.itemcouponexpiredate2.focus();
            return;
        }

        if (frmcoupon.itemcouponstartdate.value>frmcoupon.itemcouponexpiredate.value) {
            alert('���� �Ⱓ�� �߸��ƽ��ϴ�. �Ⱓ�� Ȯ�����ּ���.');
            frmcoupon.itemcouponexpiredate.focus();
            return;
        }

        if (frmcoupon.margintype.value.length<1){
            alert('���� ������ ������ �ּ���.');
            frmcoupon.margintype.focus();
            return;
        }

        if (frmcoupon.margintype.value==20){
            if (frmcoupon.defaultmargin.value.length<1){
                alert('������ �Է��� �ּ���.');
                frmcoupon.defaultmargin.focus();
                return;
            }
        }

        if (isEditMode){
            if (confirm('���� �Ͻðڽ��ϱ�?')){
                save_image().then(function(data){
                    frmcoupon.submit();
                });
            }
        }else{
            if (confirm('���� �Ͻðڽ��ϱ�?')){
                frmcoupon.submit();
            }
        }
    }

    function updateImage(isDel){
        frmcoupon.mode.value="imageupload";
        if (confirm(isDel ? '���� �Ͻðڽ��ϱ�?' : '���� �Ͻðڽ��ϱ�?')){
            save_image().then(function(data){
                frmcoupon.submit();
            });
        }
    }

    function fnSwTimeCp(obj) {
        if(obj.checked) {
            document.frmcoupon.itemcouponexpiredate2.readOnly = false;
            document.frmcoupon.itemcouponexpiredate2.className = "text";

        } else {
            document.frmcoupon.itemcouponexpiredate2.readOnly = true;
            document.frmcoupon.itemcouponexpiredate2.className = "text_ro";
            document.frmcoupon.itemcouponexpiredate2.value = "23:59:59";
        }
    }

    $(document).ready(function(){
        $("#couponImageFile").change(function(event){
            const file = event.target.files[0];

            if (!file.type.match("image.*")) {
                $("#couponImageFile").val("");
                return alert("only image");
            }

            let reader = new FileReader();
            reader.readAsDataURL(file);

            reader.onload = function(e){
                $("#couponImage").attr("src", e.target.result);
                $("#couponImageDiv").css("display", "block");
                $("#delete_image_button").css("display", "inline");
                $("#update_image_button").css("display", "inline");
            }

            changeCouponImageFlag = true;
        });

        <% IF oitemcouponmaster.FOneItem.Fitemcouponimage <> "" THEN %>
            $("#couponImageDiv").css("display", "block");
            $("#delete_image_button").css("display", "block");
            $("#update_image_button").css("display", "block");
        <% END IF %>
    });

    function delete_image(){
        $("#couponImageFile").val("");
        $("#itemcouponimage").val("");
        $("#couponImageDiv").css("display", "none");
        $("#delete_image_button").css("display", "none");
        $("#update_image_button").css("display", "none");
        updateImage(true);
    }

    function save_image(){
        return new Promise(function (resolve, reject) {
            if("<%= itemcouponidx %>" == "0"){
                alert("�����ڵ尡 �����ϴ�. ���� ��� �� �̹����� �÷��ּ���.");
                return reject();
            }

            if(changeCouponImageFlag){
                let api_url;
                if (location.hostname.startsWith('webadmin')) {
                    api_url = 'https://upload.10x10.co.kr';
                } else {
                    api_url = 'http://testupload.10x10.co.kr';
                }

                const imgData = new FormData();
                imgData.append('coupon_image', document.getElementById("couponImageFile").files[0]);
                imgData.append("coupon_code", "<%= itemcouponidx %>");
                imgData.append("reg_year", "<%= LEFT(oitemcouponmaster.FOneItem.Fregdate, 4) %>");

                $.ajax({
                    url: api_url + "/linkweb/coupon/coupon_admin_imgreg_json.asp"
                    , type: "POST"
                    , processData: false
                    , contentType: false
                    , data: imgData
                    , crossDomain: true
                    , success: function (data) {
                        //console.log(data);
                        const response = JSON.parse(data);

                        $("input[name=itemcouponimage]").val(response.coupon_image);

                        return resolve();
                    }
                    , error : function (request,status,error){
                        console.log("code", request.status);
                        console.log("message", request.responseText);
                        console.log("error", error);

                        return reject();
                    }
                });
            }else{
                return resolve();
            }
        });
    }
</script>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="a">
	<form name="frm" method="get" action="">
	<input type="hidden" name="page" value="1">
   	<tr height="10" valign="bottom" bgcolor="F4F4F4">
	        <td width="10" align="right"><img src="/images/tbl_blue_round_01.gif" width="10" height="10"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td background="/images/tbl_blue_round_02.gif"></td>
	        <td width="10" align="left" ><img src="/images/tbl_blue_round_03.gif" width="10" height="10"></td>
	</tr>
	<tr height="25" valign="bottom" bgcolor="F4F4F4">
	        <td background="/images/tbl_blue_round_04.gif"></td>
	        <td valign="top" bgcolor="F4F4F4">
	        	������ȣ : <input type="text" name="itemcouponidx" value="<%= itemcouponidx %>" Maxlength="12" size="12" readonly >
	        </td>
	        <td valign="top" align="right" bgcolor="F4F4F4">
	        	<input type="image" src="/admin/images/search2.gif" width="74" height="22" border="0"></a>
	        </td>
	        <td background="/images/tbl_blue_round_05.gif"></td>
	</tr>
	</form>
</table>

<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="a" bgcolor=#BABABA>
<form name=frmcoupon method=post action="itemcoupon_Process.asp">
<input type=hidden name="itemcouponidx" value="<%= itemcouponidx %>">
<input type=hidden name="mode" value="couponmaster">

<tr bgcolor="#DDDDFF">
	<td width="100">������</td>
	<td bgcolor="#FFFFFF"><input type="text" class="text" name="itemcouponname" value="<%= oitemcouponmaster.FOneItem.Fitemcouponname %>" size="40" maxlength="30"></td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >��������</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="itemcouponexplain" value="<%= oitemcouponmaster.FOneItem.Fitemcouponexplain %>" size="60" maxlength="50">
		<br><b><font color="red">�� �������� ����Ʈ�� ����Ǵ� �����Ͽ� �ۼ��� �ּ���.</font></b><b><font color="red"></font></b>
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >���� �̹���(����)</td>
	<td bgcolor="#FFFFFF">
	    <%
	        IF itemcouponidx > 0 THEN
	    %>
            <input type="file" id="couponImageFile" value="" />
            <input id="delete_image_button" type="button" value="����" onclick="delete_image()" style="display: none;"/>
            <input id="update_image_button" type="button" value="����" onclick="updateImage(false)" style="display: none;"/>
            <div id="couponImageDiv" class="thumbnail-area" style="display: none;">
                <img id="couponImage" src="<%=oitemcouponmaster.FOneItem.FitemcouponimageUrl%>" class="img-thumbnail link" style="width:200px;max-height:50%;" />
            </div>
        <%
            ELSE
        %>
            <b><font color="red">�� �����̹����� ��� �� ���� ���������� ��� �����մϴ�.</font></b>
        <%
            END IF
        %>

        <input type="text" name="itemcouponimage" id="itemcouponimage" value="<%=oitemcouponmaster.FOneItem.Fitemcouponimage%>" style="display: none"/>
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td width="100">��������</td>
	<td bgcolor="#FFFFFF">
	    <input type="radio" name="couponGubun" value="C" <%= ChkIIF(oitemcouponmaster.FOneItem.FcouponGubun="C","checked","") %> >�Ϲ�
	    <input type="radio" name="couponGubun" value="T" <%= ChkIIF(oitemcouponmaster.FOneItem.FcouponGubun="T","checked","") %> >Ÿ��(E-mailƯ��)
	    <input type="radio" name="couponGubun" value="P" <%= ChkIIF(oitemcouponmaster.FOneItem.FcouponGubun="P","checked","") %> >�����ι߱�(����Ʈ �߱� �Ұ� : �ý����� ����)
	    <input type="radio" name="couponGubun" value="V" <%= ChkIIF(oitemcouponmaster.FOneItem.FcouponGubun="V","checked","") %> >���̹���������(����Ʈ �߱� �Ұ�)
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >������</td>
	<td bgcolor="#FFFFFF">
		<input type="text" class="text" name="itemcouponvalue" value="<%= oitemcouponmaster.FOneItem.Fitemcouponvalue %>" size="6">
		<input type="radio" name="itemcoupontype" value="1" <% if oitemcouponmaster.FOneItem.Fitemcoupontype="1" then response.write "checked" %> > %
		<input type="radio" name="itemcoupontype" value="2" <% if oitemcouponmaster.FOneItem.Fitemcoupontype="2" then response.write "checked" %> > ��
		<input type="radio" name="itemcoupontype" value="3" <% if oitemcouponmaster.FOneItem.Fitemcoupontype="3" then response.write "checked" %> > ��۷��������� (<%=Cstr(getDefaultBeasongPayByDate(now()))%> �Է�)
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >����Ⱓ</td>
	<td bgcolor="#FFFFFF">
	<input type="text" class="text" name="itemcouponstartdate" value="<%= Left(oitemcouponmaster.FOneItem.Fitemcouponstartdate,10) %>" size="10" maxlength="10">
	<input type="text" class="text_ro" readonly name="itemcouponstartdate2" value="<%= ChkIIF(oitemcouponmaster.FOneItem.Fitemcouponstartdate<>"",Right(oitemcouponmaster.FOneItem.Fitemcouponstartdate,8),"00:00:00") %>" size="8" maxlength="8" />
	<a href="javascript:calendarOpen(frmcoupon.itemcouponstartdate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
	~
	<input type="text" class="text" name="itemcouponexpiredate" value="<%= Left(oitemcouponmaster.FOneItem.Fitemcouponexpiredate,10) %>" size="10" maxlength="10">
	<input type="text" <%= ChkIIF(oitemcouponmaster.FOneItem.Fcoupontype="T","class=""text""","class=""text_ro"" readonly")%> name="itemcouponexpiredate2" value="<%= ChkIIF(oitemcouponmaster.FOneItem.Fitemcouponexpiredate<>"",Right(oitemcouponmaster.FOneItem.Fitemcouponexpiredate,8),"23:59:59") %>" size="8" maxlength="8" />
	<a href="javascript:calendarOpen(frmcoupon.itemcouponexpiredate);"><img src="/images/calicon.gif" border="0" align="absmiddle" height=21></a>
	/ <input type="checkbox" name="coupontype" value="T" <%= ChkIIF(oitemcouponmaster.FOneItem.Fcoupontype="T","checked","")%> class="checkbox" onclick="fnSwTimeCp(this)" />������������
	<br>(<%= Left(now(),10) %> 00:00:00)  ~  (<%= Left(now(),10) %> 23:59:59)
	<br><font color="#808080">(�� ���� �̹� �ٿ�ε��� ������ ����Ⱓ�� ������� �ʽ��ϴ�. ���� �Ⱓ �����ÿ� �������ּ���.)</font>
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >�⺻ ��������</td>
	<td bgcolor="#FFFFFF">
		<select name="margintype" onchange="AlertMarginChange();fninput(this.value);">
		<!--<option value="">---����--- -->
		<option value="30" <% if oitemcouponmaster.FOneItem.Fmargintype="30" then response.write "selected" %> >���ϸ���
		<option value="60" <% if oitemcouponmaster.FOneItem.Fmargintype="60" then response.write "selected" %> >��ü�δ�
		<option value="50" <% if oitemcouponmaster.FOneItem.Fmargintype="50" then response.write "selected" %> >�ݹݺδ�
		<option value="10" <% if oitemcouponmaster.FOneItem.Fmargintype="10" then response.write "selected" %> >�ٹ����ٺδ�
		<option value="20" <% if oitemcouponmaster.FOneItem.Fmargintype="20" then response.write "selected" %> >��������
		<option value="00" <% if oitemcouponmaster.FOneItem.Fmargintype="00" then response.write "selected" %> >��ǰ��������
		<option value="90" <% if oitemcouponmaster.FOneItem.Fmargintype="90" then response.write "selected" %> >20%��ü���
		<option value="80" <% if oitemcouponmaster.FOneItem.Fmargintype="80" then response.write "selected" %> >������(500��ü�δ�)
		</select>
		<span id="marginlayer" style="display:<% IF oitemcouponmaster.FOneItem.Fmargintype<>"20" Then response.write "none" %>"><input type="text" class="text" name="defaultmargin" value="<%=oitemcouponmaster.FOneItem.FDefaultMargin%>" size="3" maxlength="3" onChange="AlertMarginChange();">%</span>
		<font color="#808080">(��ǰ���� �������� �ٸ� ��� ������ ���� �����մϴ�.)</font>
	</td>
</tr>
<% if oitemcouponmaster.FOneItem.FisuedCount>0  then %>
<tr bgcolor="#DDDDFF">
	<td >�߱�������</td>
	<td bgcolor="#FFFFFF" style="color:#E03333; font-weight:bold;"><%= FormatNumber(oitemcouponmaster.FOneItem.FisuedCount,0) %></td>
</tr>
<% end if %>
<tr bgcolor="#DDDDFF">
	<td >�߱� ����</td>
	<td bgcolor="#FFFFFF">
	<%= oitemcouponmaster.FOneItem.GetOpenStateName %>
	<% if (oitemcouponmaster.FOneItem.FItemCouponIdx>0) then %>
    	<% if (oitemcouponmaster.FOneItem.IsOpenAvailCoupon) then %>
    	--&gt;<input type="button" value="����" onclick="OpenCouponMaster();" <% if (itemcouponidx="1056") or (itemcouponidx="1200") or (itemcouponidx="1202") or (itemcouponidx="1385") or (itemcouponidx="1716") or (itemcouponidx="2174")  then response.write "disabled" %> >
    	<% elseif (oitemcouponmaster.FOneItem.Fopenstate="0")  then %>
    	--&gt;<input type="button" value="�߱޿���" onclick="reserveCouponMaster();" <% if (itemcouponidx="2768")then response.write "disabled" %> >
    	<% elseif (oitemcouponmaster.FOneItem.Fopenstate="9")  then %>

    	<% else %>
    	--&gt;<input type="button" value="�߱ް�������" onclick="CloseCouponMaster();" <% if (oitemcouponmaster.FOneItem.Fapplyitemcount>1000) or (itemcouponidx="4487") or (itemcouponidx="3543") or (itemcouponidx="1202") or (itemcouponidx="1385") or (itemcouponidx="1716") or (itemcouponidx="2174") or (itemcouponidx="2768")then response.write "disabled" %> >
    	(������ 12�� 15�п� �ڵ� ����˴ϴ�.)
    	<% end if %>
    <% end if %>
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >�����</td>
	<td bgcolor="#FFFFFF">
		<%= oitemcouponmaster.FOneItem.Fregdate %>
	</td>
</tr>
<tr bgcolor="#DDDDFF">
	<td >����������</td>
	<td bgcolor="#FFFFFF">
		<%= oitemcouponmaster.FOneItem.FlastupDt %>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<% if (IsEditMode) then %>
	    <% if (oitemcouponmaster.FOneItem.Fopenstate="0") then %>
	    <td colspan="2" align="center"><input type="button" value="�� ��" onclick="SaveCouponMaster(frmcoupon, true)"></td>
	    <% elseif (Not oitemcouponmaster.FOneItem.IsOpenAvailCoupon) then %>
	    <td colspan="2" align="center"><input type="button" value="�� ��" onclick="SaveCouponMaster(frmcoupon, true)"  Disabled></td>
	    <% else %>
	    <td colspan="2" align="center"><input type="button" value="�� ��" onclick="SaveCouponMaster(frmcoupon, true)"></td>
	    <% end if %>
	<% else %>
	<td colspan="2" align="center"><input type="button" value="�� ��" onclick="SaveCouponMaster(frmcoupon, false)"></td>
	<% end if %>
</tr>
</form>
</table>

<%
set oitemcouponmaster = Nothing
%>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->