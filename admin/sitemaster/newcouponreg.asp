<%@ language=vbscript %>
<% option explicit %>
<%
Response.AddHeader "Cache-Control","no-cache"
Response.AddHeader "Expires","0"
Response.AddHeader "Pragma","no-cache"
%>
<%
'####################################################
' Description :  ���ʽ� ����
' History : ������ ����
'			2022.07.04 �ѿ�� ����(isms���������)
'####################################################
%>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->
<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/adminbodyhead.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/classes/items/newcouponcls.asp" -->
<%
dim dispCate : dispCate = requestCheckvar(request("dispCate"),32)
dim idx, ocoupon
	idx = requestCheckvar(request("idx"),10)
	if idx="" then idx=0

set ocoupon = new CCouponMaster
	ocoupon.FRectIdx = idx

	if idx<>0 then
		ocoupon.GetOneCouponMaster ''GetCouponMasterList
	else
		set ocoupon.FOneItem = new CCouponMasterItem
	end if
%>

<script type="text/javascript" src="https://cdn.jsdelivr.net/npm/es6-promise@4/dist/es6-promise.min.js"></script>
<script type="text/javascript" src="https://cdn.jsdelivr.net/npm/es6-promise@4/dist/es6-promise.auto.min.js"></script>
<script language="JavaScript" src="/js/jquery-1.7.1.min.js"></script>
<script lanuage='javscript'>
    let changeCouponImageFlag = false;

    $(function(){
        showTip();
         $("[name=isopenlistcoupon]").click(function(){
            showTip();
         })
    })

    function showTip(){
        if(document.frm.isopenlistcoupon.value == "N"){
            $("#tip").css("display","");
        }else{
            $("#tip").css("display","none");
        }
    }

    function submitForm(frm){
        if (frm.couponname.value.length<1){
            alert('�������� �Է��ϼ���.');
            frm.couponname.focus();
            return;
        }

        //���������� üũ
        if (frm.isfreebeasongcoupon.checked){
    //        if ((!frm.coupontype[1].checked)||(frm.couponvalue.value*1!=2000)||(frm.minbuyprice.value*1!=0)){
    //            alert('������ ������ ��� ����Ÿ�� - �� ���αݾ� 2000, �ּұ��űݾ� 0���� ���� �˴ϴ�.');
    //
    //            frm.coupontype[1].checked = true;
    //            frm.couponvalue.value = 2000;
    //            frm.minbuyprice.value = 0;
    //            return;
    //        }
        }else{
            if ((!frm.coupontype[0].checked)&&(!frm.coupontype[1].checked)){
                alert('���� Ÿ���� �����ϼ���.');
                frm.coupontype[0].focus();
                return;
            }

            if (frm.couponvalue.value.length<1){
                alert('���� �ݾ��̳� %�� �Է��ϼ���.');
                frm.couponvalue.focus();
                return;
            }

            if (frm.minbuyprice.value.length<1){
                alert('�ּ� ���űݾ��� �Է��ϼ���.');
                frm.minbuyprice.focus();
                return;
            }
        }

        //ī�װ�/�귣�������� ��� %���� ���úҰ� (����.)
        //2019-11-13 MD���� �������̵��� �̺�Ʈ �������� ���� �귣��,ī�װ� % ���� ������� ����
        if ((frm.targetcpntype.value=="B")||(frm.targetcpntype.value=="C")){
            if (frm.isfreebeasongcoupon.checked){
                alert('ī�װ�,�귣�� ������ ������ ������ ������ �� �����ϴ�.');
                return;
            }

            if (frm.coupontype[0].checked){
                if (!confirm('ī�װ�,�귣�� % ���������� ��ǰ������ ���� ������ �ȵ� �� �ֽ��ϴ�.\n�����Ͻðڽ��ϱ�?')){
                    return;
                }
                <% ' ����Ʈ�� �÷��������ϰ�� �귣������ ���������� ����� ����. ���� ���� ����. %>
                alert('ī�װ�,�귣�� ������ % ������ ������ �� �����ϴ�. �ݾ������� ����ϼ���.');
                return;
            }
        }
        if ((frm.targetcpntype.value=="B")){
            if (frm.brandShareValue.value>50){
                alert('�귣�������� ��ü �д����� 50%�� ������ �����ϴ�.');
                return;
            }
        }
        if ((frm.coupontype[0].checked)&&(frm.mxCpnDiscount.value.length<1)){
            alert('�ִ����αݾ��� �Է��ϼ���.');
            frm.startdate.focus();
            return;
        }
        if (frm.startdate.value.length<1){
            alert('��ȿ�Ⱓ �������� �Է��ϼ���.');
            frm.startdate.focus();
            return;
        }

        if (frm.expiredate.value.length<1){
            alert('��ȿ�Ⱓ �������� �Է��ϼ���.');
            frm.expiredate.focus();
            return;
        }

        if (frm.openfinishdate.value.length<1){
            alert('���� �߱� �������� �Է��ϼ���.');
            frm.openfinishdate.focus();
            return;
        }

        if ((frm.validsitename.value=="academy")||(frm.validsitename.value=="diyitem")){
            if (frm.isfreebeasongcoupon.checked){
                alert('�ΰŽ� ��ī���� �����ΰ�� ������ ���� üũ �Ұ�');
                frm.isfreebeasongcoupon.focus();
                return;
            }

            if (!confirm('�ΰŽ� ��ī���� �������� �����ϼ̽��ϴ�. ����Ͻðڽ��ϱ�?')){
                return;
            }
        }

        var ret = confirm('���� �Ͻðڽ��ϱ�?');

        if (ret){
            save_image().then(function(data){
               frm.submit();
            });
        }
    }

    function EnableBox(comp){
        if (comp.checked){
            frm.targetitemlist.disabled = false;
            frm.couponmeaipprice.disabled = false;

            frm.targetitemlist.style.backgroundColor = "#FFFFFF";
            frm.couponmeaipprice.style.backgroundColor = "#FFFFFF";
        }else{
            frm.targetitemlist.disabled = true;
            frm.couponmeaipprice.disabled = true;

            frm.targetitemlist.style.backgroundColor = "#E6E6E6";
            frm.couponmeaipprice.style.backgroundColor = "#E6E6E6";
        }

    }

    function disableType(comp){
        var frm = comp.form;
        if (comp.name=="isfreebeasongcoupon"){
            frm.couponvalue.disabled = comp.checked;
            frm.coupontype[0].disabled = comp.checked;
            frm.coupontype[1].disabled = comp.checked;
            //frm.minbuyprice.disabled = comp.checked;
            frm.mxCpnDiscount.disabled = comp.checked;

        }else if (comp.name=="targetcpntype"){
            if (comp.value=="C"){
                document.getElementById("brandSBtn").style.display = "none";
                document.getElementById("cateSelBtn").style.display = "block";
                frm.isfreebeasongcoupon.disabled = true;
            }else if (comp.value=="B"){
                document.getElementById("brandSBtn").style.display = "block";
                document.getElementById("cateSelBtn").style.display = "none";
                frm.isfreebeasongcoupon.disabled = true;
            }else{
                document.getElementById("brandSBtn").style.display = "none";
                document.getElementById("cateSelBtn").style.display = "none";
                frm.isfreebeasongcoupon.disabled = false;
            }
        }
        chkCpnType(frm);
    }

    function jsSearchDispCate(frmname,targetcompname, targetcpndtlnm){
        var dispCate = eval(frmname+'.'+targetcompname).value;
        var uri = '/common/module/popDispCateSelect.asp?dispCate='+dispCate+'&frmname='+frmname+'&targetcompname='+targetcompname+'&targetcpndtlnm='+targetcpndtlnm;
        var popwin = window.open(uri,'popDispCateSelect','width=800, height=400, scrollbars=yes, resizable=yes');
        popwin.focus();
    }

    function chkCpnType(o){
        var dctp = o.coupontype;
        var tgtp = o.targetcpntype;
        if (dctp.value=="1"&&tgtp.value==""){
            document.getElementById("imxcpndiscount_tr").style.display = "";
        }else{
            frm.mxCpnDiscount.value=0;
            document.getElementById("imxcpndiscount_tr").style.display = "none";
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
            }

            changeCouponImageFlag = true;
        });

        <% IF ocoupon.FOneItem.Fcouponimage <> "" THEN %>
            $("#couponImageDiv").css("display", "block");
            $("#delete_image_button").css("display", "block");
        <% END IF %>
    });

    function delete_image(){
        $("#couponImageFile").val("");
        $("#usercouponimage").val("");
        $("#couponImageDiv").css("display", "none");
        $("#delete_image_button").css("display", "none");
    }

    function save_image(){
        return new Promise(function (resolve, reject) {
            if("<%= ocoupon.FOneItem.FIdx %>" == "0"){
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
                imgData.append("coupon_code", "<%= ocoupon.FOneItem.FIdx %>");
                imgData.append("reg_year", "<%= LEFT(ocoupon.FOneItem.Fregdate, 4) %>");

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

                        $("input[name=usercouponimage]").val(response.coupon_image);

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

<form name="frm" method="post" action="/admin/sitemaster/docoupon.asp">
<input type="hidden" name="menupos" value="<%= menupos %>">
<input type="hidden" name="idx" value="<%= ocoupon.FOneItem.FIdx %>">
<table width="900" border="0" cellpadding="3" cellspacing="1" bgcolor=#3d3d3d class=a>
<tr>
	<td bgcolor="#DDDDFF" width="120">IDx</td>
	<td bgcolor="#FFFFFF"><%= ocoupon.FOneItem.FIdx %></td>
</tr>
<tr>
	<td bgcolor="#DDDDFF">������</td>
	<td bgcolor="#FFFFFF"><input type=text name=couponname value="<%= ocoupon.FOneItem.Fcouponname %>" maxlength="100" size=80>
	<br>(ex �ٹ����� �ָ� ����)</td>
</tr>
<tr>
	<td bgcolor="#DDDDFF">���� �̹���(����)</td>
	<td bgcolor="#FFFFFF">
        <%
            IF ocoupon.FOneItem.FIdx > 0 THEN
        %>
            <input type="file" id="couponImageFile" value="" />
            <input id="delete_image_button" type="button" value="����" onclick="delete_image()" style="display: none;"/>
            <div id="couponImageDiv" class="thumbnail-area" style="display: none;">
                <img id="couponImage" src="<%=ocoupon.FOneItem.FcouponimageUrl%>" class="img-thumbnail link" style="width:200px;max-height:50%;" />
            </div>
        <%
            ELSE
        %>
            <b><font color="red">�� �����̹����� ��� �� ���� ���������� ��� �����մϴ�.</font></b>
        <%
            END IF
        %>

        <input type="text" name="usercouponimage" id="usercouponimage" value="<%=ocoupon.FOneItem.Fcouponimage%>" style="display: none"/>
    </td>
</tr>
<!-- 2018/01/18 �߰� -->
<tr>
	<td bgcolor="#DDDDFF">����Ÿ��II</td>
	<td bgcolor="#FFFFFF">
	    <label style="margin-right:5px;"><input type="radio" name="targetcpntype" value=""  <%= CHKIIF(ocoupon.FOneItem.Ftargetcpntype="","checked","") %> onClick="disableType(this);">�Ϲ�</label>
	    <label style="margin-right:5px;"><input type="radio" name="targetcpntype" value="B" <%= CHKIIF(ocoupon.FOneItem.Ftargetcpntype="B","checked","") %> onClick="disableType(this);" >�귣������</label>
	    <label style="margin-right:6px;"><input type="radio" name="targetcpntype" value="C" <%= CHKIIF(ocoupon.FOneItem.Ftargetcpntype="C","checked","") %> onClick="disableType(this);" >ī�װ�����</label>

	    <div id="brandSBtn" <%= CHKIIF(ocoupon.FOneItem.Ftargetcpntype="B","style='display:block'","style='display:none'") %> >
	        <p>��� �귣�� :
			<input type="text" name="targetcpnsourcebrand" value="<%=ocoupon.FOneItem.Ftargetcpnsource%>" size="20" maxlength="32" readonly >
	        <input type="button" class="button" value="�귣��˻�" onclick="jsSearchBrandID(this.form.name,'targetcpnsourcebrand');">
			</p>
			<p>��ü �д��� :
				<input type="text" name="brandShareValue" value="<%=chkIIF(ocoupon.FOneItem.FbrandShareValue="","0",ocoupon.FOneItem.FbrandShareValue)%>" size="3" style="text-align:right;" /> %
			</p>
	    </div>
	    <div id="cateSelBtn" <%= CHKIIF(ocoupon.FOneItem.Ftargetcpntype="C","style='display:block'","style='display:none'") %> >
	        <input type="text" name="targetcpnsourcecate" value="<%=ocoupon.FOneItem.Ftargetcpnsource%>" size="20" maxlength="32" readonly >
	        <input type="text" name="targetcpndtlnm" value="<%=ocoupon.FOneItem.getTargetCateName%>" size="40"  readonly>
	        <input type="button" class="button" value="ī�װ�����" onclick="jsSearchDispCate(this.form.name,'targetcpnsourcecate','targetcpndtlnm');" >
	    </div>
	</td>
</tr>
<tr>
	<td bgcolor="#DDDDFF">����������<br>(�ٹ����ٹ��)</td>
	<td bgcolor="#FFFFFF">
		<%
		''//[Fingers]����Ʈ����>>���ʽ��������θ�� ������������ �ΰŽ� ������ �ٹ����� ���� ��� ����. '���� all ����
		if menupos <> "1224" and menupos <> "1216" then
		%>
	    	<input type="checkbox" name="isfreebeasongcoupon" value="Y" <% if ocoupon.FOneItem.IsFreedeliverCoupon then response.write "checked" %> onClick="disableType(this);"> ����������
	    <% else %>
	    	<input type="checkbox" name="isfreebeasongcoupon" value="Y" disabled <% if ocoupon.FOneItem.IsFreedeliverCoupon then response.write "checked" %> onClick="disableType(this);"> ����������
	    <% end if %>
	    <!--
	    <br>
	    <input type="checkbox" name="isweekendcoupon" value="Y" <% if ocoupon.FOneItem.IsWeekendCoupon then response.write "checked" %> > �ָ� ����
        -->
	</td>
</tr>

<tr>
	<td bgcolor="#DDDDFF">����Ÿ��</td>
	<td bgcolor="#FFFFFF">
		<input type=text name=couponvalue value="<%= ocoupon.FOneItem.Fcouponvalue %>" maxlength=7 size=10 <% if ocoupon.FOneItem.IsFreedeliverCoupon then response.write "disabled" %> >
	    <label style="margin-right:5px;"><input type="radio" name="coupontype" value="1" <%=chkIIF(ocoupon.FOneItem.IsFreedeliverCoupon,"disabled","")%> <%=chkIIF(ocoupon.FOneItem.Fcoupontype="1","checked","")%> onClick="chkCpnType(this.form)" />%����</label>
	    <label style="margin-right:5px;"><input type="radio" name="coupontype" value="2" <%=chkIIF(ocoupon.FOneItem.IsFreedeliverCoupon,"disabled","")%> <%=chkIIF(ocoupon.FOneItem.Fcoupontype="2" or ocoupon.FOneItem.Fcoupontype="","checked","")%> onClick="chkCpnType(this.form)" />������</label>
		(�ݾ� �Ǵ� % ����)
	</td>
</tr>
<!--
<% if (FALSE) then %>
<tr>
	<td bgcolor="#DDDDFF" width="100">Ư����ǰ����</td>
	<% if ocoupon.FOneItem.IsTargetItemCoupon then %>
		<td bgcolor="#FFFFFF">
		Ư����ǰ ���� �����: <input type=checkbox name=targetitemusing onclick="EnableBox(this)" checked ><br>
		��ǰ��ȣ: <input type=text name=targetitemlist value="<%= ocoupon.FOneItem.Ftargetitemlist %>" size=9 maxlength=9  >(Ư�� ��ǰ�� ���ε�)
		&nbsp;&nbsp;
		��������� ���԰�: <input type=text name=couponmeaipprice value="<%= ocoupon.FOneItem.Fcouponmeaipprice %>" size=7 maxlength=9  >(��ü�δ��� ��� ���԰� ����)
		</td>
	<% else %>
		<td bgcolor="#FFFFFF">
		Ư����ǰ ���� �����: <input type=checkbox name=targetitemusing onclick="EnableBox(this)"><br>
		��ǰ��ȣ: <input type=text name=targetitemlist value="<%= ocoupon.FOneItem.Ftargetitemlist %>" size=9 maxlength=9 disabled style='background-color:#E6E6E6;'>(Ư�� ��ǰ�� ���ε�)
		&nbsp;&nbsp;
		��������� ���԰�: <input type=text name=couponmeaipprice value="<%= ocoupon.FOneItem.Fcouponmeaipprice %>" size=7 maxlength=9 disabled style='background-color:#E6E6E6;'>(��ü�δ��� ��� ���԰� ����)
		</td>
	<% end if %>
</tr>
<% end if %>
-->
<tr>
	<td bgcolor="#DDDDFF">�ּұ��űݾ�</td>
	<td bgcolor="#FFFFFF"><input type="text" name="minbuyprice" value="<%= ocoupon.FOneItem.Fminbuyprice %>" maxlength="7" size="10" />�� �̻� ���Ž� ��밡��(����)</td>
</tr>
<tr id="imxcpndiscount_tr" <%=CHKIIF((ocoupon.FOneItem.Fcoupontype="1" and ocoupon.FOneItem.Ftargetcpntype="") or ocoupon.FOneItem.FmxCpnDiscount>0,"style='display:'","style='display:none'")%>>
	<td bgcolor="#DDDDFF">�ִ����αݾ�</td>
	<td bgcolor="#FFFFFF"><input type="text" name="mxCpnDiscount" value="<%= ocoupon.FOneItem.FmxCpnDiscount %>" maxlength="7" size="10" />�� ����(����)(ex 5% �� 10000 / 10%�� 20000 / ������ 0 �Է�)</td>
</tr>

<tr>
	<td bgcolor="#DDDDFF">��ȿ�Ⱓ</td>
	<td bgcolor="#FFFFFF">
	    <input type=text name=startdate value="<%= ocoupon.FOneItem.Fstartdate %>" maxlength=19 size=20>~<input type=text name=expiredate value="<%= ocoupon.FOneItem.Fexpiredate %>" maxlength=19 size=20>
	    (<%= Left(now(),10) %> 00:00:00 ~ <%= Left(now(),10) %> 23:59:59)
	</td>
</tr>
<tr>
	<td bgcolor="#DDDDFF">�����߱޸�����</td>
	<td bgcolor="#FFFFFF"><input type=text name="openfinishdate" value="<%= ocoupon.FOneItem.Fopenfinishdate %>" maxlength=19 size=20>(2004-04-31 23:59:59)</td>
</tr>
<tr>
	<td bgcolor="#DDDDFF">���ó</td>
	<td bgcolor="#FFFFFF">
		<label style="margin-right:5px;"><input type="radio" name="validsitename" value="" <%= CHKIIF(ocoupon.FOneItem.Fvalidsitename="","checked","") %> >�ٹ����� ���ʽ� ����</label>
		<!-- ����
		<label><input type="radio" name="validsitename" value="academy" <%= CHKIIF(ocoupon.FOneItem.Fvalidsitename="academy","checked","") %> >�ΰŽ� ��ī���� ���� ����</label>
		<label><input type="radio" name="validsitename" value="diyitem" <%= CHKIIF(ocoupon.FOneItem.Fvalidsitename="diyitem","checked","") %> >�ΰŽ� ��ī���� ��ǰ ����</label>
		-->
		<label style="margin-right:5px;"><input type="radio" name="validsitename" value="mobile" <%= CHKIIF(ocoupon.FOneItem.Fvalidsitename="mobile","checked","") %> >����� ���ʽ� ����</label>
		<label style="margin-right:5px;"><input type="radio" name="validsitename" value="app" <%= CHKIIF(ocoupon.FOneItem.Fvalidsitename="app","checked","") %> >APP ���ʽ� ����</label>
	</td>
</tr>
<tr>
	<td bgcolor="#DDDDFF">��Ÿ�ڸ�Ʈ</td>
	<td bgcolor="#FFFFFF"><textarea name="etcstr" cols=80 rows=8><%= ReplaceBracket(ocoupon.FOneItem.Fetcstr) %></textarea></td>
</tr>
<tr>
	<td bgcolor="#DDDDFF">��ü��������</td>
	<td bgcolor="#FFFFFF">
		<label style="margin-right:5px;"><input type="radio" name="isopenlistcoupon" value="N" <% if ocoupon.FOneItem.Fisopenlistcoupon="N" then Response.Write "checked" %> />��ü��</label>
		<label style="margin-right:5px;"><input type="radio" name="isopenlistcoupon" value="Y" <% if ocoupon.FOneItem.Fisopenlistcoupon="Y" or ocoupon.FOneItem.Fisopenlistcoupon="" then Response.Write "checked" %> />���ð�(������)</label>
		<label style="margin-right:5px;"><input type="radio" name="isopenlistcoupon" value="J" <% if ocoupon.FOneItem.Fisopenlistcoupon="J" then Response.Write "checked" %> />ȸ����������</label>
		<label style="margin-right:5px;"><input type="radio" name="isopenlistcoupon" value="M" <% if ocoupon.FOneItem.Fisopenlistcoupon="M" then Response.Write "checked" %> />����� �����</label>
		<div id="tip" style="color:red;display:none">**��ü�� ���ý� �ش� ������ ��ȿ�Ⱓ �� �α��ν� ������ �߱޵˴ϴ�.</div>
	</td>
</tr>
<tr>
	<td bgcolor="#DDDDFF">��뿩��</td>
	<td bgcolor="#FFFFFF">
		<label style="margin-right:5px;"><input type="radio" name="isusing" value="Y" <%=chkIIF(ocoupon.FOneItem.FIsUsing="Y","checked","")%> />Y</label>
		<label style="margin-right:5px;"><input type="radio" name="isusing" value="N" <%=chkIIF(ocoupon.FOneItem.FIsUsing="N","checked","")%> />N</label>
	</td>
</tr>
<tr>
	<td colspan="2" align=center bgcolor="#FFFFFF"><input type=button value="����" onClick="submitForm(frm);" class="button"></td>
</tr>
</table>
</form>
<%
set ocoupon = Nothing
%>
<!-- #include virtual="/admin/lib/adminbodytail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->