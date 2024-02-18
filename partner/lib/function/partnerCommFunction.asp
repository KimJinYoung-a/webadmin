<%
'20161109 ¼öÁ¤
Sub ptDrawDateBox(byval yyyy1,mm1,dd1,yyyy2,mm2,dd2, frm)
%>
<script type="text/javascript">
function jsResetValidDate(yyyy, mm, dd, svalue) { 
		 
	var year = yyyy;
  var month = mm;
	var day = dd;

    var lastdate = new Array(31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31);

	if(((0 == (year % 4)) && (0 != (year % 100))) || (0 == (year % 400))) {
		lastdate[1] = 29;
	}

 	if (day >= (lastdate[month-1]+1)) { 
		eval("document.all."+svalue).value = lastdate[month-1]; 
	} 
	 
}
</script>
<%
	dim buf,i

	buf = "<select class='formSlt' name='yyyy1'>"
    for i=Year(now) to 2003 STEP -1
		if (CStr(i)=CStr(yyyy1)) then
			buf = buf + "<option value='" + CStr(i) +"' selected>" + CStr(i) + "</option>"
		else
    		buf = buf + "<option value=" + CStr(i) + " >" + CStr(i) + "</option>"
		end if
	next
    buf = buf + "</select>"

    buf = buf + "<select class='formSlt' name='mm1' onchange=""jsResetValidDate("+frm+".yyyy1.value, "+frm+".mm1.value, "+frm+".dd1.value,'dd1');"">"
    for i=1 to 12
		if (Format00(2,i)=Format00(2,mm1)) then
			buf = buf + "<option value='" + Format00(2,i) +"' selected>" + Format00(2,i) + "</option>"
		else
    	    buf = buf + "<option value='" + Format00(2,i) +"' >" + Format00(2,i) + "</option>"
		end if
	next

    buf = buf + "</select>"

    buf = buf + "<select class='formSlt' name='dd1' onchange=""jsResetValidDate("+frm+".yyyy1.value, "+frm+".mm1.value, "+frm+".dd1.value,'dd1');"">"

    for i=1 to 31
		if (Format00(2,i)=Format00(2,dd1)) then
	    buf = buf + "<option value='" + Format00(2,i) +"' selected>" + Format00(2,i) + "</option>"
		else
        buf = buf + "<option value='" + Format00(2,i) + "' >" + Format00(2,i) + "</option>"
		end if
    next
    buf = buf + "</select>"

    buf = buf + "~"

    buf = buf + "<select class='formSlt' name='yyyy2'>"
    for i=Year(now)+1 to 2003 step -1
		if (CStr(i)=CStr(yyyy2)) then
			buf = buf + "<option value='" + CStr(i) +"' selected>" + CStr(i) + "</option>"
		else
    		buf = buf + "<option value=" + CStr(i) + " >" + CStr(i) + "</option>"
		end if
	next
    buf = buf + "</select>"

    buf = buf + "<select class='formSlt' name='mm2' onchange=""jsResetValidDate("+frm+".yyyy2.value, "+frm+".mm2.value, "+frm+".dd2.value,'dd2');"">"
    for i=1 to 12
		if (Format00(2,i)=Format00(2,mm2)) then
			buf = buf + "<option value='" + Format00(2,i) +"' selected>" + Format00(2,i) + "</option>"
		else
    	    buf = buf + "<option value='" + Format00(2,i) +"' >" + Format00(2,i) + "</option>"
		end if
	next

    buf = buf + "</select>"

    buf = buf + "<select class='formSlt' name='dd2' onchange=""jsResetValidDate("+frm+".yyyy2.value, "+frm+".mm2.value, "+frm+".dd2.value,'dd2');"">"
    for i=1 to 31
		if (Format00(2,i)=Format00(2,dd2)) then
			buf = buf + "<option value='" + Format00(2,i) +"' selected>" + Format00(2,i) + "</option>"
		else
    	    buf = buf + "<option value='" + Format00(2,i) +"' >" + Format00(2,i) + "</option>"
		end if
    next
    buf = buf + "</select>"

    response.write buf
end Sub


Sub ptDrawYMBox(byval yyyy1,mm1)
	dim buf,i

	buf = "<select class='formSlt' name='yyyy1'>"
    for i=Year(now)+1 to 2003 step -1
		if (CStr(i)=CStr(yyyy1)) then
			buf = buf + "<option value='" + CStr(i) +"' selected>" + CStr(i) + "</option>"
		else
    	buf = buf + "<option value=" + CStr(i) + " >" + CStr(i) + "</option>"
        end if
	next
    buf = buf + "</select>"

    buf = buf + "<select class='formSlt' name='mm1' >"

    for i=1 to 12
		if (Format00(2,i)=Format00(2,mm1)) then
			buf = buf + "<option value='" + Format00(2,i) +"' selected>" + Format00(2,i) + "</option>"
		else
    	    buf = buf + "<option value='" + Format00(2,i) +"' >" + Format00(2,i) + "</option>"
		end if
	next

    buf = buf + "</select>"

    response.write buf
end Sub
%>