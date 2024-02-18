<%@ language=vbscript %>
<% option explicit %>
<%
'###########################################################
' Description :  업체정보
' History : 2007.10.26 한용민 수정
'###########################################################
%>
<!-- #include virtual="/cscenterv2/lib/incSessionAdminCS.asp" -->
<!-- #include virtual="/cscenterv2/lib/db/dbopen.asp" -->
<!-- #include virtual="/cscenterv2/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/cscenterv2/lib/classes/user/partnerusercls.asp"-->
<!-- #include virtual="/cscenterv2/lib/classes/cs/cs_aslistcls.asp"-->

<script language='javascript'>
window.resizeTo(600,700);
</script>
<%
dim ogroup,opartner,i
dim makerid , takbae
dim groupid

makerid = RequestCheckvar(request("makerid"),32)
takbae = RequestCheckvar(request("takbaebox"),16)

set opartner = new CPartnerUser
opartner.FRectDesignerID = makerid
opartner.GetOnePartnerNUser


set ogroup = new CPartnerGroup

if opartner.FResultCount>0 then
	ogroup.FRectGroupid = opartner.FOneItem.FGroupid
	ogroup.GetOneGroupInfo
end if


dim OReturnAddr
set OReturnAddr = new CCSReturnAddress

OReturnAddr.FRectMakerid = makerid
OReturnAddr.GetBrandReturnAddress


dim OCSBrandMemo
set OCSBrandMemo = new CCSBrandMemo

OCSBrandMemo.FRectMakerid = makerid
OCSBrandMemo.GetBrandMemo

dim brandmemo_found
if (OCSBrandMemo.Fbrandid = "") then
	brandmemo_found = "N"
else
	brandmemo_found = "Y"
end if


%>
<script language="javascript">

function SaveBrandInfo(frm){
	var ret = confirm('저장 하시겠습니까?');

	if (ret){
		frm.submit();
	}
}

function jsPopCal(fName,sName)
{
	var fd = eval("document."+fName+"."+sName);

	if(fd.readOnly==false)
	{
		var winCal;
		winCal = window.open('/lib/common_cal.asp?FN='+fName+'&DN='+sName,'pCal','width=250, height=200');
		winCal.focus();
	}
}
</script>

<table width="100%" align="center" cellpadding="3" cellspacing="1" class="a" bgcolor="<%= adminColor("tablebg") %>">
	<tr height="25" bgcolor="<%= adminColor("tabletop") %>">
		<td colspan="4">
			<b>브랜드 정보</b>
		</td>
	</tr>

	<tr>
		<td colspan="4" bgcolor="#FFFFFF" height="25">[ 브랜드 기본정보 ] (동일한 업체라도 브랜드별로 반품정보가 다를 수 있습니다.)</td>
	</tr>

	<tr height="25">
		<td width="18%" bgcolor="<%= adminColor("tabletop") %>" >브랜드ID</td>
		<td width="40%" bgcolor="#FFFFFF"><b><%= opartner.FOneItem.FID %></b></td>
		<td width="18%" bgcolor="<%= adminColor("tabletop") %>">스트리트명</td>
		<td bgcolor="#FFFFFF"><b><%= opartner.FOneItem.Fsocname_kor %></b></td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">반품담당자</td>
		<td bgcolor="#FFFFFF"><%= OReturnAddr.FreturnName %></td>
		<td bgcolor="<%= adminColor("tabletop") %>">반품전화</td>
		<td bgcolor="#FFFFFF"><%= OReturnAddr.FreturnPhone %></td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">반품핸드폰</td>
		<td bgcolor="#FFFFFF"><%= OReturnAddr.Freturnhp %></td>
		<td bgcolor="<%= adminColor("tabletop") %>">반품이메일</td>
		<td bgcolor="#FFFFFF"><%= OReturnAddr.FreturnEmail %></td>
	</tr>

	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">반품 주소</td>
		<td colspan="3" bgcolor="#FFFFFF" >
			[<%= OReturnAddr.FreturnZipcode %>] <%= OReturnAddr.FreturnZipaddr %> <%= OReturnAddr.FreturnEtcaddr %>
		</td>
	</tr>

	<tr height="25">
		<td colspan="4" bgcolor="#FFFFFF" height="25">[ 브랜드 배송정보 ]</td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">조건배송여부</td>
		<td bgcolor="#FFFFFF">
			<% if (opartner.FOneItem.IsFreeBeasong) then %>
				항상 무료배송
			<% end if %>
			<% if (opartner.FOneItem.IsUpcheReceivePayDeliverItem) then %>
				착불배송
			<% end if %>
			<% if opartner.FOneItem.IsUpcheParticleDeliverItem then %>
				가격별 무료배송
			<% end if %>
			<% if ((opartner.FOneItem.IsUpcheParticleDeliverItem) or (opartner.FOneItem.IsUpcheReceivePayDeliverItem)) and Not(opartner.FOneItem.IsFreeBeasong) then %>
			<% else %>
				N
			<% end if %>
		</td>
		<td bgcolor="<%= adminColor("tabletop") %>">배송조건</td>
		<td bgcolor="#FFFFFF">
			<% if opartner.FOneItem.IsUpcheParticleDeliverItem then %>
			<b><%=FormatNumber(opartner.FOneItem.FdefaultFreeBeasongLimit,0)%></b>원 이상 구매시 무료<br>
			배송비 <b><%=FormatNumber(opartner.FOneItem.FdefaultDeliverPay,0)%></b>원
			<% end if %>
		</td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">거래택배사</td>
		<td bgcolor="#FFFFFF" colspan="3"><%= opartner.FOneItem.Ftakbae_name %> (<%= opartner.FOneItem.Ftakbae_tel %>)</td>
	</tr>

	<tr height="25">
		<td colspan="4" bgcolor="#FFFFFF" height="25">[ 브랜드 추가정보 ]</td>
	</tr>
	<tr height="25">
		<form name=brandmemo method=post action="do_brandmemo_input.asp">
		<input type=hidden name=makerid value="<%= makerid %>">
		<input type=hidden name=mode value="<% if brandmemo_found = "Y" then %>modify<% else %>insert<% end if %>">
		<td bgcolor="<%= adminColor("tabletop") %>">회수가능여부</td>
		<td bgcolor="#FFFFFF" colspan="3">
			<select class="select" name="is_return_allow">
		     	<option value='-' >-</option>
		     	<option value='Y' <% if (OCSBrandMemo.Fis_return_allow = "Y") then %>selected<% end if %>>Y</option>
		     	<option value='N' <% if (OCSBrandMemo.Fis_return_allow = "N") then %>selected<% end if %>>N</option>
	     	</select>
	    </td>
	</tr>

	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">상담가능시간</td>
		<td bgcolor="#FFFFFF" colspan="3">
			<select class="select" name="tel_start">
				<option value='0'>-- : --</option>
		     	<% for i = 6 to 15 %>
		     	<option value='<%= i %>' <% if (OCSBrandMemo.Ftel_start = i) then %>selected<% end if %>><%= i %>:00</option>
		    	<% next %>
	     	</select>
	     	~
			<select class="select" name="tel_end">
				<option value='0'>-- : --</option>
		     	<% for i = 12 to 21 %>
		     	<option value='<%= i %>' <% if (OCSBrandMemo.Ftel_end = i) then %>selected<% end if %>><%= i %>:00</option>
		    	<% next %>
	     	</select>
	      	(토요일 근무여부
	      	<select class="select" name="is_saturday_work">
		     	<option value='-' >-</option>
		     	<option value='Y' <% if (OCSBrandMemo.Fis_saturday_work = "Y") then %>selected<% end if %>>Y</option>
		     	<option value='N' <% if (OCSBrandMemo.Fis_saturday_work = "N") then %>selected<% end if %>>N</option>
	     	</select>)
	     </td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">휴가일정</td>
		<td bgcolor="#FFFFFF" colspan="3">
			<input type="text" size="10" name="vacation_startday" value="<%= OCSBrandMemo.Fvacation_startday %>" onClick="jsPopCal('brandmemo','vacation_startday');" style="cursor:hand;"> - <input type="text" size="10" name="vacation_endday" value="<%= OCSBrandMemo.Fvacation_endday %>" onClick="jsPopCal('brandmemo','vacation_endday');" style="cursor:hand;">
	     </td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">기타메모</td>
		<td bgcolor="#FFFFFF" colspan="3">
		<textarea class="textarea" name=brand_comment cols="70" rows="5"><% if (OCSBrandMemo.Fbrand_comment = "") then %>각종메모(비상연락망,환불계좌,맞교환가능여부 등)<% else %><%= OCSBrandMemo.Fbrand_comment %><% end if %></textarea>
		</td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">최종수정일</td>
		<td bgcolor="#FFFFFF" colspan="3">
		<% if Len(OCSBrandMemo.Flast_modifyday) > 10 then %>
			<%= Left(OCSBrandMemo.Flast_modifyday) %>
		<% else %>
			<%= (OCSBrandMemo.Flast_modifyday) %>
		<% end if %>
		</td>
	</tr>
	<tr height="25" align="center">
		<td colspan="4" bgcolor="#FFFFFF" height="25">
			<input type="button" class="button" value="추가정보수정" onclick="SaveBrandInfo(brandmemo)"></td>
		</td>
	</tr>
	</form>


	<tr height="25">
		<td colspan="4" bgcolor="#FFFFFF" height="25">[ 업체기본정보 ]</td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">회사명(상호)</td>
		<td bgcolor="#FFFFFF"><b><%= ogroup.FOneItem.FCompany_name %></b></td>
		<td bgcolor="<%= adminColor("tabletop") %>">그룹코드</td>
		<td bgcolor="#FFFFFF"><b><%= opartner.FOneItem.FGroupid %></b></td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">대표전화</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fcompany_tel %></td>
		<td bgcolor="<%= adminColor("tabletop") %>">팩스</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fcompany_fax %></td>
	</tr height="25">
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">사무실 주소</td>
		<td bgcolor="#FFFFFF" colspan=3>[<%= ogroup.FOneItem.Freturn_zipcode %>] <%= ogroup.FOneItem.Freturn_address %> <%= ogroup.FOneItem.Freturn_address2 %></td>
	</tr height="25">



	<tr>
		<td colspan="4" bgcolor="#FFFFFF" height="25">[ 업체 담당자정보 ]</td>
	</tr>

	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">담당자명</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fmanager_name %></td>
		<td bgcolor="<%= adminColor("tabletop") %>">일반전화</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fmanager_phone %></td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">E-Mail</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fmanager_email %></td>
		<td bgcolor="<%= adminColor("tabletop") %>">핸드폰</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fmanager_hp %></td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">배송담당자명</td>
		<td bgcolor="#FFFFFF" colspan="3">브랜드별로 조회 가능합니다</td>
	</tr>
	<!--
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">배송담당자명</td>
		<td bgcolor="#FFFFFF" colspan="3"><%= ogroup.FOneItem.Fdeliver_name %></td>
		<td bgcolor="<%= adminColor("tabletop") %>">일반전화</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fdeliver_phone %></td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">E-Mail</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fdeliver_email %></td>
		<td bgcolor="<%= adminColor("tabletop") %>">핸드폰</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fdeliver_hp %></td>
	</tr>
	-->
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">정산담당자명</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fjungsan_name %></td>
		<td bgcolor="<%= adminColor("tabletop") %>">일반전화</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fjungsan_phone %></td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">E-Mail</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fjungsan_email %></td>
		<td bgcolor="<%= adminColor("tabletop") %>">핸드폰</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fjungsan_hp %></td>
	</tr>














	<tr height="25">
		<td colspan="4" bgcolor="#FFFFFF" height="25">[ 업체 사업자등록정보 ]</td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">회사명(상호)</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.FCompany_name %></td>
		<td bgcolor="<%= adminColor("tabletop") %>">대표자</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fceoname %></td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">사업자번호</td>
		<td bgcolor="#FFFFFF" colspan="3"><%= ogroup.FOneItem.Fcompany_no %></td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">사업장소재지</td>
		<td colspan="3" bgcolor="#FFFFFF" >
			<%= ogroup.FOneItem.Fcompany_zipcode %>&nbsp;
			<%= ogroup.FOneItem.Fcompany_address %>&nbsp;
			<%= ogroup.FOneItem.Fcompany_address2 %>
		</td>
	</tr>
	<tr height="25">
		<td bgcolor="<%= adminColor("tabletop") %>">업태</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fcompany_uptae %></td>
		<td bgcolor="<%= adminColor("tabletop") %>">업종</td>
		<td bgcolor="#FFFFFF"><%= ogroup.FOneItem.Fcompany_upjong %></td>
	</tr>


	<tr align="center">
		<td colspan="4" bgcolor="#FFFFFF" height="25">
			<input type="button" class="button" value="닫기" onclick="self.close();"></td>
		</td>
	</tr>

</table>

<%
set opartner = Nothing
set ogroup = Nothing
%>
<!-- #include virtual="/common/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->