<%option explicit%>
<html>
<head>
<script>
function show_() //not used
{
var obj = document.getElementsByName("chk");
var tmp='';
//alert(obj.length);
for (i=0;i<obj.length;i++)
{
if (obj[i].checked==true)
{
tmp+=obj[i].value+',';
}
}
//alert('src_kpi_detail.asp?project='+tmp+'&yearmonth='+document.getElementById("yearmonth").value);

parent.frames['src_kpi_detail'].location.href = 'src_kpi_throughput.asp?project='+tmp+'&yearmonth='+document.getElementById("yearmonth").value;
//getJiraItems via so_reporting
}

function show2_()
{
var obj = document.getElementsByName("chk");
var tmp='';
//alert(obj.length);
for (i=0;i<obj.length;i++)
{
if (obj[i].checked==true)
{
tmp+=obj[i].value+',';
}
}
//alert(tmp);
//return false;
//alert('src_kpi_detail.asp?project='+tmp+'&yearmonth='+document.getElementById("yearmonth").value);
parent.frames['src_kpi_detail'].location.href = 'src_kpi_ontime_missrate_done.asp?project='+tmp+'&yearmonth='+document.getElementById("yearmonth").value;
//getJiraItems via so_reporting
}

function show3_() //not used
{
var obj = document.getElementsByName("chk");
var tmp='';
//alert(obj.length);
for (i=0;i<obj.length;i++)
{
if (obj[i].checked==true)
{
tmp+=obj[i].value+',';
}
}
//alert('src_kpi_detail.asp?project='+tmp+'&yearmonth='+document.getElementById("yearmonth").value);
parent.frames['src_kpi_detail'].location.href = 'src_kpi_ontime_missrate.asp?project='+tmp+'&yearmonth='+document.getElementById("yearmonth").value;
}

function show4_()
{
var obj = document.getElementsByName("chk");
var tmp='';
//alert(obj.length);
for (i=0;i<obj.length;i++)
{
if (obj[i].checked==true)
{
tmp+=obj[i].value+',';
}
}
//alert('src_kpi_detail.asp?project='+tmp+'&yearmonth='+document.getElementById("yearmonth").value);
parent.frames['src_kpi_detail'].location.href = 'src_kpi_fulfill_missrate.asp?project='+tmp+'&yearmonth='+document.getElementById("yearmonth").value;
//getJiraItems via so_reporting
}

function show5_()
{
var obj = document.getElementsByName("chk");
var tmp='';
//alert(obj.length);
for (i=0;i<obj.length;i++)
{
if (obj[i].checked==true)
{
tmp+=obj[i].value+',';
}
}
//alert('baseline_enddate_shifts.asp?project='+tmp+'&yearmonth='+document.getElementById("yearmonth").value);
parent.frames['src_kpi_detail'].location.href = 'baseline_enddate_shifts.asp?project='+tmp+'&yearmonth='+document.getElementById("yearmonth").value;
//getJiraItems via so_reporting
}

function show_sor_risk_impact_reduction_()
{
parent.frames['src_kpi_detail'].location.href = 'show_sor_risk_impact_reduction.asp?yearmonth='+document.getElementById("yearmonth").value;
}

function show_sor_risk_decision_deadline_missrate_()
{
parent.frames['src_kpi_detail'].location.href = 'show_sor_risk_decision_deadline_missrate.asp?yearmonth='+document.getElementById("yearmonth").value;
}

function show_soim_adherence_to_4_10_deadlines_rc_()
{
parent.frames['src_kpi_detail'].location.href = 'show_soim_adherence_to_4_10_deadlines_rc.asp?yearmonth='+document.getElementById("yearmonth").value;
}

function show_soim_adherence_to_4_10_deadlines_ac_()
{
parent.frames['src_kpi_detail'].location.href = 'show_soim_adherence_to_4_10_deadlines_ac.asp?yearmonth='+document.getElementById("yearmonth").value;
}

</script>
</head>
<body style='font-family:verdana;font-size:8pt;'>
<b>Filters</b><br>
<br>
<!--
Choose project<br>
<input type='checkbox' name='chk' id='chk' value='STSPES'>STSPES<br>
<input         type='checkbox' name='chk' id='chk' value='STSCNT'>STSCNT<br>
<input         type='checkbox' name='chk' id='chk' value='MOEBIUS'>MOEBIUS<br>
<input         type='checkbox' name='chk' id='chk' value='RES'>RES<br>
<input checked         type='checkbox' name='chk' id='chk' value='STS'>STS<br>
<br>
-->
<%
'response.Write "Choose period<br>"
'response.Write "<select  style='font-family:verdana;font-size:8pt;' name='yearmonth' id='yearmonth'>"
dim y
dim m
dim dt1
dim dt2
dim dt

dt1 = dateadd("m",5,date)
dt2 = dateadd("m",-7,dt1)
'response.write dt1 & " - " & dt2
response.Write "Choose period<br>"
response.Write "<select  style='font-family:verdana;font-size:8pt;' name='yearmonth' id='yearmonth'>"
dt = dt2
while dt <= dt1
	y = year(dt)
	m = month(dt)
	if year(date) = y and month(date) = m then
	response.Write "<option selected value='" & right("0000"&y,4) & right("00"&m,2) & "'>" & monthname(m) & " " & right("0000"&y,4) & "</option>"
	else
	response.Write "<option value='" & right("0000"&y,4) & right("00"&m,2) & "'>" & monthname(m) & " " & right("0000"&y,4) & "</option>"
	end if
	dt = dateadd("m", 1, dt)
wend
'for y = year(date)+5 to year(date)-5 step -1
'for m = 12 to 1 step -1
'	if year(date) = y and month(date) = m then
'	response.Write "<option selected value='" & right("0000"&y,4) & right("00"&m,2) & "'>" & monthname(m) & " " & right("0000"&y,4) & "</option>"
'	else
'	response.Write "<option value='" & right("0000"&y,4) & right("00"&m,2) & "'>" & monthname(m) & " " & right("0000"&y,4) & "</option>"
'	end if
'next
'next
response.Write "</select><br>"
%>
<%
'<br>
'Choose 'In Progress' states<br>
'dim arr
'dim arr2
'dim a
'
'arr = split("Open||0@@In Progress||1@@Development||1@@Development Done||1@@Validation||1@@Validation Tracking||1@@Done||0@@", "@@")
'for a = lbound(arr) to ubound(arr)
'if arr(a)<> ""then
'	arr2=split(arr(a),"||")
'	if arr2(1) = "1" then
'	response.Write "<input type='checkbox' checked name='inprogress' id='inprogress' value='" & arr2(0) & "'>" & arr2(0) & "<br>"
'	else
'	response.Write "<input type='checkbox'         name='inprogress' id='inprogress' value='" & arr2(0) & "'>" & arr2(0) & "<br>"
'	end if
'end if
'next
'<br>
%>
<hr noshade>
<!--<input type='button' value='Throughput' onclick='show_();' style='font-family:verdana;font-size:8pt;'><br>-->
<!--<input type='button' disabled value='Rejection rate'         onclick='show2_();' style='font-family:verdana;font-size:8pt;'><br>-->
<!--<input type='button'  value='On time miss rate'      onclick='show3_();' style='font-family:verdana;font-size:8pt;'><br>-->
<!--<input type='button'  value='On time miss rate'      onclick='show2_();' style='font-family:verdana;font-size:8pt;'><br>-->
<input type='button'  value='Fulfillment miss rate'      onclick='show4_();' style='font-family:verdana;font-size:8pt;'><br>
<!--<input type='button'  value='Baseline Enddate shifts'      onclick='show5_();' style='font-family:verdana;font-size:8pt;'><br>-->
<hr noshade>
<input type='button'  value='SOR Risk Impact Reduction' 			onclick='show_sor_risk_impact_reduction_();' style='font-family:verdana;font-size:8pt;'><br>
<input type='button'  value='SOR Risk Decision Deadline Miss Rate'	onclick='show_sor_risk_decision_deadline_missrate_();' style='font-family:verdana;font-size:8pt;'><br>
<input type='button'  value='SOIM Adherence to 14 deadline Root Cause'	onclick='show_soim_adherence_to_4_10_deadlines_rc_();' style='font-family:verdana;font-size:8pt;'><br>
<input type='button'  value='SOIM Adherence to 14 deadline Actions'		onclick='show_soim_adherence_to_4_10_deadlines_ac_();' style='font-family:verdana;font-size:8pt;'><br>
<br>
</body>
</html>