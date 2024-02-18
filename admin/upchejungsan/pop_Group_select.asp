<%@ language=vbscript %>
<% option explicit %>
<!-- #include virtual="/admin/incSessionAdmin.asp" -->

<!-- #include virtual="/lib/db/dbopen.asp" -->
<!-- #include virtual="/admin/lib/popheader.asp"-->
<!-- #include virtual="/lib/util/htmllib.asp"-->
<!-- #include virtual="/lib/function.asp"-->
<!-- #include virtual="/admin/upchejungsan/upchejungsan_function.asp"-->
<%
Dim ipFileNo : ipFileNo=requestCheckVar(request("ipFileNo"),10)
Dim bank : bank=requestCheckVar(request("bank"),10)
Dim bankaccount : bankaccount=requestCheckVar(request("bankaccount"),20)
Dim firstSel : firstSel=requestCheckVar(request("idx"),10)
Dim sqlStr, arrDetailList, intLoop

arrDetailList = fnGetJFixIpkumList(ipFileNo)

function getJGetStateName(jstate)
    if IsNULL(jstate) then
        getJGetStateName="미지정"
        exit function
    end if

    jstate = CStr(jstate)
    if jstate="0" then
		getJGetStateName = "수정중"
	elseif jstate="1" then
	    getJGetStateName = "업체확인대기"
	elseif jstate="2" then
	    getJGetStateName = "업체확인완료"
	elseif jstate="3" then
		getJGetStateName = "정산확정"
	elseif jstate="7" then
		getJGetStateName = "입금완료"
	else
        getJGetStateName = jstate
	end if
end function

function getJGetStateColor(jstate)
    if IsNULL(jstate) then
        getJGetStateColor="#FF0000"
        exit function
    end if

    jstate=CStr(jstate)
    if jstate="0" then
		getJGetStateColor = "#000000"
	elseif jstate="1" then
	    getJGetStateColor = "#448888"
	elseif jstate="2" then
	    getJGetStateColor = "#0000FF"
	elseif jstate="3" then
		getJGetStateColor = "#0000FF"
	elseif jstate="7" then
		getJGetStateColor = "#FF0000"
	else

	end if
end function
%>
<script>
function SelectItems(){	
var itemcount = 0;
var frm;
var ck=0;
frm = document.frm;

if(typeof(frm.chkitem) !="undefined"){
    if(!frm.chkitem.length){
        if(!frm.chkitem.checked){
            alert("선택한 상품이 없습니다. 상품을 선택해 주세요");
            return;
        }
            frm.itemidarr.value = frm.chkitem.value;
    }else{
        for(i=0;i<frm.chkitem.length;i++){
            if(frm.chkitem[i].checked) {
                ck=ck+1;	   	    			
                if (frm.itemidarr.value==""){
                    frm.itemidarr.value =  frm.chkitem[i].value;
                }else{
                    frm.itemidarr.value = frm.itemidarr.value + "," +frm.chkitem[i].value;
                } 
            }	
        }
        
        if (frm.itemidarr.value == ""){
            alert("선택한 상품이 없습니다. 상품을 선택해 주세요");
            return;
        }
    }
}else{
    alert("추가할 상품이 없습니다.");
    return;
}
	frm.action = "dobankingupflag.asp";
	frm.submit();
}
function jsChkAll(){	
var frm;
frm = document.frm;
	if (frm.chkAll.checked){			      
	   if(typeof(frm.chkitem) !="undefined"){
	   	   if(!frm.chkitem.length){
		   	 	frm.chkitem.checked = true;	   	 
		   }else{
				for(i=0;i<frm.chkitem.length;i++){
					frm.chkitem[i].checked = true;
			 	}		
		   }	
	   }	
	} else {	  
	  if(typeof(frm.chkitem) !="undefined"){
	  	if(!frm.chkitem.length){
	   	 	frm.chkitem.checked = false;	  
	   	}else{
			for(i=0;i<frm.chkitem.length;i++){
				frm.chkitem[i].checked = false;
			}	
		}		
	  }	
	
	}
	
}
</script>
<% IF isArray(arrDetailList) THEN %>
<form name="frm" method="post">
<input type="hidden" name="itemidarr">
<input type="hidden" name="ipFileNo" value="<%=ipFileNo%>">
<input type="hidden" name="firstSel" value="<%=firstSel%>">
<input type="hidden" name="mode" value="ipkumGroupMulti">
<table width="100%" border="0" align="center" class="a" cellpadding="3" cellspacing="1" bgcolor="<%= adminColor("tablebg") %>">
    <tr align="center" bgcolor="<%= adminColor("tabletop") %>">
		<td width="120">브랜드ID</td>
		<td width="60">상태</td>
		<td width="60">은행</td>
		<td width="80">계좌</td>
		<td width="80">정산금액</td>
		<td>업체명</td>
		<td width="50">그룹코드</td>
		<td width="50">Erp코드</td>
		<td width="30">선택<input type="checkbox" name="chkAll" onClick="jsChkAll();"></td>
	</tr>
	<%  For intLoop = 0 To UBound(arrDetailList,2) %>
    <% if (trim(bank)=trim(arrDetailList(9,intLoop))) and (trim(bankaccount)=trim(arrDetailList(10,intLoop))) then %>
	<tr align="center" bgcolor="#FFFFFF">
		<td><%= arrDetailList(8,intLoop) %></td>
		<td><font color="<%= getJGetStateColor(arrDetailList(12,intLoop)) %>"><%= getJGetStateName(arrDetailList(12,intLoop)) %></font></td>
		<td><%= arrDetailList(9,intLoop) %></td>
		<td><%= arrDetailList(10,intLoop) %></td>
		<td align="right">
		<% if arrDetailList(4,intLoop)<1 then %><font color=red><% else %><font color="#000000"><% end if %>
		<% if Not isNULL(arrDetailList(4,intLoop)) then %><%= FormatNumber(arrDetailList(4,intLoop),0) %><% end if %>
		</font>
		</td>
		<td><%= arrDetailList(16,intLoop) %></td>
		<td><%= arrDetailList(7,intLoop) %></td>
		<td <%=CHKIIF(arrDetailList(18,intLoop)=0 or isNULL(arrDetailList(19,intLoop)),"bgcolor='#CCCCCC'","") %> >
		    <% if IsNULL(arrDetailList(19,intLoop)) then %>
		    <% else %>
		    <%= arrDetailList(19,intLoop) %>
		    <% end if %>
		</td>
		<td>
            <% if IsNULL(arrDetailList(14,intLoop)) or arrDetailList(14,intLoop)="" then %>
                <input type="checkbox" name="chkitem" value="<%= arrDetailList(0,intLoop) %>">
            <% end if %>
        </td>
	</tr>
    <% end if %>
	<%  next %>
	<tr bgcolor="#FFFFFF">
		<td colspan="9" align="right"><input type="button" value="그룹묶기" class="button" onClick="SelectItems();"></td>
	</tr>
</table>
</form>
<% end if %>
<!-- #include virtual="/admin/lib/poptail.asp"-->
<!-- #include virtual="/lib/db/dbclose.asp" -->