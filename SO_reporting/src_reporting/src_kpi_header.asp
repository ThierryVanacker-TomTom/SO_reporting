<%option explicit%>
<html>
<head>
<script>
function show_sts_inprogressthroughput_()
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
parent.frames['src_kpi_detail'].location.href = 'src_kpi_inprogress_throughput.asp?project='+tmp+'&yearmonth='+document.getElementById("yearmonth").value;
}

function show_sts_rejection_()
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
parent.frames['src_kpi_detail'].location.href = 'src_kpi_rejection_rate.asp?project='+tmp+'&yearmonth='+document.getElementById("yearmonth").value;
}

function show_sts_ontimemissrate_()
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

function show_sso_rejection_()
{
parent.frames['src_kpi_detail'].location.href = 'src_kpi_sso_rejection_rate.asp?yearmonth='+document.getElementById("yearmonth").value;
}

function show_sso_ontimemissrate_()
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
parent.frames['src_kpi_detail'].location.href = 'src_kpi_sso_ontime_missrate.asp?project='+tmp+'&yearmonth='+document.getElementById("yearmonth").value;
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
dim y
dim m
dim dt1
dim dt2
dim dt

dt1 = date
dt2 = dateadd("yyyy",-1,dt1)
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
'for y = year(date) to year(date)-5 step -1
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
STS<br>
<!-- show all items -->
<!-- <input type='button' value='In Progress throughput' onclick='show_sts_inprogressthroughput_();' style='font-family:verdana;font-size:8pt;'><br> -->
<!-- show all items that did not make it from Validation straint to rejected -->
<input type='button' value='Rejection rate'         onclick='show_sts_rejection_();' style='font-family:verdana;font-size:8pt;'><br>
<!-- show all items that were intended to finish in this month vs the target to finish in this month-->
<input type='button' value='On time miss rate'      onclick='show_sts_ontimemissrate_();' style='font-family:verdana;font-size:8pt;'><br>
<!--
<br>
<i>The output will show all items of the selected project(s) where the state is set to Done for the selected period</i>
-->
<hr noshade>
SSO - Engineering<br>
<input type='button' value='Rejection rate'    onclick='show_sso_rejection_();'      style='font-family:verdana;font-size:8pt;'><br>
<input type='button' value='On time miss rate' onclick='show_sso_ontimemissrate_();' style='font-family:verdana;font-size:8pt;'><br>

</body>
</html>