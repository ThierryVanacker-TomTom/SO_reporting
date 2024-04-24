<%option explicit%>
<%server.scripttimeout=60*60 '60minutes%>
<!--#include file="jsonObject.class.asp" -->
<html>
<head>
<script>
function downloadCSV(tbl,filename)
{
	//loop through table rows & cols
	var obj = document.getElementById(tbl);
	var row;
	var col;
	var csv='';
	for (var i = 0; row = obj.rows[i]; i++)
	{
	   //iterate through rows
	   //rows would be accessed using the "row" variable assigned in the for loop
	   for (var j = 0; col = row.cells[j]; j++)
	   {
	     //iterate through columns
	     //columns would be accessed using the "col" variable assigned in the for loop
	     var str = col.innerText;
	     str = str.replace("#", encodeURIComponent("#"));
	     csv+='"'+str+'";';
	   }
	   csv+='\n';
	}
	filename = filename;
	if (!csv.match(/^data:text\/csv/i))
	{
	csv = 'data:text/csv;charset=utf-8,' + encodeURI('\uFEFF'+csv); //force the file to be UTF8
	}
	var data = csv;
	//alert(data);
	var link = document.createElement('a');
	link.setAttribute('href', data);
	link.setAttribute('download', filename);
	document.body.appendChild(link); //needed for FF
	link.click();
}
</script>
</head>
<body style='font-family:verdana;font-size:8pt;'>
<%
Dim tmp
Dim tmp2
Dim url
dim id
dim prj
'dim arr
dim issuetype
dim event_
dim sum_
dim desc_
dim i
dim j

dim arr
dim a
dim arr2
dim a2
dim arrP
dim aP

dim xml
dim objXML
set objXML = createobject("MSXML.DOMDocument")
Dim item_
Dim field_in_item_
dim nodelist

dim xml2
dim objXML2
set objXML2 = createobject("MSXML.DOMDocument")
Dim item2_
Dim field_in_item2_
dim nodelist2

dim label
dim hist_
dim comp
dim bln
dim bln2

dim duedate
dim target_duedate
dim sprint
dim sprintenddate

dim yearmonth
yearmonth = request.querystring("yearmonth")
dim settodone

dim created
dim decision_created_on
dim subtasks
dim parent
dim parentcreated
dim processed_parents
processed_parents = "||"

dim dt1
dim dt2

dim fso
dim f

dim dt
dim from_
dim to_

yearmonth = request.querystring("yearmonth")
if (yearmonth = "") then yearmonth = right("0000" & year(date),4) & right("00" & Month(date),2)

'kata: changed dateadd from "w", -10 to extend the deadline to 14 calendar days
tmp = dateserial(mid(yearmonth,1,4), mid(yearmonth,5,2), 1)
'format into 2018/06/30
dt = dateadd("d", -14, tmp)
from_ = right("0000" & year(dt),4) & "/" & right("00" & month(dt),2) & "/" & right("00" & day(dt),2)
dt = dateadd("m", 1, tmp)
dt = dateadd("d", -14, dt)
to_ = right("0000" & year(dt),4) & "/" & right("00" & month(dt),2) & "/" & right("00" & day(dt),2)

dim jql
jql = ""
jql = jql & " project = ""SOIM"" "
jql = jql & " and issuetype in (""Prevent Recurrence"",""Corrective action"") " ',""Fix""
jql = jql & " and issueFunction in subtasksOf("" "
jql = jql & " project = 'SOIM' "
'14JUN2022 - added 4 new issuetypes : Service"",""Tool"",""Program"",""Process
jql = jql & " AND issuetype in ('Source Incident','MoMa Incident','KPI','Audit', 'Service','Tool','Program','Process') "
jql = jql & " AND createdDate >= '" & from_ & "' AND createdDate <= '" & to_ & "' "") "

'jql = jql & " AND ""Assigned Unit"" in (""SO MOMA"",""SO SSO"",""SO APA"",""SO AME"",""SO PMO"",""SO GDT"",""SO EAP"", ""SO ECA"", ""SO SAMEA"", ""SO EECA"", ""SO AFR"", ""SO WCE"", ""SO STS"", ""SO SAM"", ""SO PDV"", ""SO OCE"", ""SO NEA"", ""SO NAM"", ""SO LAM"")"
'jql = jql & " AND "
'jql = jql & " status changed to closed during (""" & from_ & """, """ & to_ & """)"
'' jql = jql & " and issuekey = 'OM-19434' "
'jql = jql & " AND ""End date"" >= 2018-01-01 "
'' jql = jql & " AND cf[20162] >= ""2018/01/01"" "

'jql = jql & " OR status changed to closed during (""" & from_ & """"", """ & to_ & """)"
'jql = jql & " OR status changed to planned during (""" & from_ & """"", """ & to_ & """)"
' jql = jql & " )"

'response.write jql
'response.end

'for testing
'jql = " issuekey = ""OM-27842"" "

xml = getJiraItems(jql)
'response.write xml
'response.end

'response.write replace(xml, "<", "[")
'response.write jql
'response.write xml
'for ap = lbound(arrp) to ubound(arrp)
'response.write arrp(ap) & "<br>"
'if arrp(ap) <> "" then
'	response.write "<br>" & now & "<br>"
'
'	prj = arrP(aP)

	'getJiraItems gets all the items from the project (no additional filtering) and turns it into an XML


'load the XML we generated before (to speed  up things)
'set fso = createobject("scripting.filesystemobject")
'set f = fso.opentextfile("c:\inetpub\wwwroot\src_reporting\src_kpi_detail.xml")
'xml = f.readall
'f.close
'set f = nothing
'set fso = nothing


	'response.write (replace(xml, "<" , "["))
	'response.write ("<hr noshade>")
	objXML.LoadXML xml

''save the the output for now (just to speed up devs)
'set fso = createobject("scripting.filesystemobject")
'set f = fso.createtextfile("c:\inetpub\wwwroot\src_reporting\src_kpi_detail.xml")
'f.writeline xml
'f.close
'set f = nothing
'set fso = nothing


	xml = ""
	xml = xml & "<item_result>" & vbcrlf
	'now rebuild the XML using addiotnal filters
	Set nodelist = objXML.getElementsByTagName("item_result/*")
	i = 1
	For Each item_ In nodelist
		bln = true
		For Each field_in_item_ In item_.ChildNodes
			' 'check if correct issuetype
			' if field_in_item_.BaseName = "issuetype" then
			' if instr("||Epic||Initiative||", "||" & field_in_item_.Text & "||") > 0 then
				' bln = false
			' end if
			' end if
			' 'check if correct components
			' if field_in_item_.BaseName = "components" then
			' arr = split(field_in_item_.Text, "||")
			' for a = lbound(arr) to ubound(arr)
			' if arr(a) <> "" then
				' if instr(lcase("||Events organization||Hosting Sharepoint sites||R&D||STS internal processes||STS Newsletter||TD platform||"),lcase("||" & arr(a) & "||")) > 0 then
					' bln = false
				' end if
			' end if
			' next
			' end if
			'check if closed in this month
			if field_in_item_.BaseName = "transitions_history" then
			    bln2 = False
			    arr = Split(field_in_item_.Text, "@@")
			    For a = LBound(arr) To UBound(arr)
			    If arr(a) <> "" Then
				arr2 = Split(arr(a), "||")
				'For a2 = LBound(arr2) To UBound(arr2)
				'If arr2(a2) <> "" Then
				    If arr2(2) = "Closed" Then
				    If Mid(arr2(0), 1, 6) = yearmonth Then
					bln2 = True
				    End If
				    End If
				'End If
				'Next
			    End If
			    Next
			    If bln2 = True Then ' it is set to closed in this month
				'do nothing
			    Else
				'not a good item
				'bln = False
			    End If
			end if

		Next
		'if bln = true then
		if true then
		xml = xml & "<item>" & vbcrlf
		For Each field_in_item_ In item_.ChildNodes
			xml = xml & "<" & field_in_item_.BaseName & ">" & xmlsafe("" & field_in_item_.Text) & "</" & field_in_item_.BaseName & ">" & vbcrlf
		Next
		xml = xml & "</item>" & vbcrlf
		end if

		i = i + 1
	Next
	xml = xml & "</item_result>" & vbcrlf
'	response.write ("<br>" & replace(xml, "<" , "["))
	objXML.LoadXML xml

	'ok now we have a good set of xml to go with
	i = 0
	j = 0
	response.write "<span style='font-family:verdana;font-size:10pt;'><b>Adherence to 4-10 deadlines - Actions</b></span><br>"
	response.write "<span style='font-family:verdana;font-size:8pt;'>"
	response.write "All issues in project SOIM and of type Source Incident, MoMa Incident, KPI or Audit which are created between " & from_ & " and " & to_ & "<br>"
	response.write "</span>"

	response.write "<table style='font-family:verdana;font-size:8pt;border-collapse:collapse;' border='1' id='table'>"
	response.write "<tr>"
	response.write "<td>" & "Pkey" & "-" & "issuenum" & "</td>"
	response.write "<td>" & "Issuetype" & "</td>"
	response.write "<td>" & "Summary" & "</td>"
	response.write "<td>" & "Assignee"   & "</td>"
	response.write "<td>" & "Created on"   & "</td>"
	response.write "<td>" & "Parent Action id" & "</td>"
	response.write "<td>" & "Parent Action issuetype" & "</td>"
	response.write "<td>" & "Parent Action created" & "</td>"
	response.write "<td>" & "Action created late?" & "</td>"
	response.write "</tr>"
	Set nodelist = objXML.getElementsByTagName("item_result/*")
	For Each item_ In nodelist
		bln = true
		'8APR2020 - quickfix if itemid = OM-65464 then skip
		'14APR2020 : not needed anymore : if getFieldValue(item_,"key") = "OM-65464" then bln = false
		if bln = true then
			'if (getFieldValue(item_,"resolution")) = "Declined"  then
			'	j = j + 1
			'	response.write "<tr style='background-color:red'>"
			'else
				response.write "<tr>"
			'end if
			response.write "<td>" & getFieldValue(item_,"key") & "</td>"
			response.write "<td>" & getFieldValue(item_,"issuetype") & "</td>"
			response.write "<td>" & getFieldValue(item_,"summary") & "</td>"
			response.write "<td>" & getFieldValue(item_,"assignee") & "</td>"
			created = getFieldValue(item_,"created")
			response.write "<td>" & toDD_MM_YYYY(created) & "</td>"

			'for each of the subtasks get the information
			parent = getFieldValue(item_,"parent")
			processed_parents = processed_parents & parent & "||"

			tmp = ""
			arr = split(parent, "||")
			for a = lbound(arr) to ubound(arr)
			if arr(a) <> "" then
				jql = "issuekey=""" & arr(a) & """"
				xml2 = getJiraItems(jql)
				objXML2.loadXML xml2
				Set nodelist2 = objXML2.getElementsByTagName("item_result/*")
				For Each item2_ In nodelist2
					tmp = tmp & arr(a) & "$$"
					tmp = tmp & getFieldValue(item2_,"issuetype") & "$$"
					tmp = tmp & getFieldValue(item2_,"created") & "##"
				next
			end if
			next
			'response.write "<td>" & tmp & "</td>"

			'at this point we extracted all the relevant info to draw conclusions
			'extract the Parent task
			arr = split(tmp, "##")
			tmp = ""
			tmp2 = ""
			for a = lbound(arr) to ubound(arr)
			if arr(a) <> "" then
				arr2 = split(arr(a), "$$")
				parent = arr2(0)
				tmp = arr2(1)
				parentcreated = arr2(2)
				'make sure to stop after first parent
				a = ubound(arr)
			end if
			next

			response.write "<td>" & parent & "</td>"
			response.write "<td>" & tmp & "</td>"
			response.write "<td>" & toDD_MM_YYYY(parentcreated) & "</td>"

 			'now put conclusion in 'action created late'
			'if parent creationdate + 10 < action created then its a yes
			'Kata: changed dateadd from "w", 10 in order to extend the deadline to 14 calendar days
			dt1 = dateserial(mid(parentcreated,1,4),mid(parentcreated,5,2),mid(parentcreated,7,2))
			dt1 = dateadd("d", 14, dt1)
			dt2 = dateserial(mid(created,1,4),mid(created,5,2),mid(created,7,2))
			if dt1 < dt2 then
				response.write "<td>" & "yes" & "</td>"
			else
				response.write "<td>" & "no" & "</td>"
			end if

			response.write "</tr>"
			i = i + 1
		end if
	next

	response.write "</table>"
	response.write "<a href='#' onclick='downloadCSV(""table"",""output.csv"");'>CSV</a><br>"
	response.write i & " items<br>" ' were set to Closed in " & yearmonth & " however " & j & " items have Quality Declined<br>"
	response.write "<br>"

'response.write processed_parents

'ALSO show SOIM items who do not have a childtaks
jql = ""
jql = jql & " project = ""SOIM"" "
jql = jql & " AND issuetype in ('Source Incident','MoMa Incident','KPI','Audit') "
jql = jql & " AND createdDate >= '" & from_ & "' AND createdDate <= '" & to_ & "' "

xml = getJiraItems(jql)
'response.write (replace(xml, "<" , "["))
objXML.LoadXML xml

	xml = ""
	xml = xml & "<item_result>" & vbcrlf
	'now rebuild the XML using addiotnal filters
	Set nodelist = objXML.getElementsByTagName("item_result/*")
	i = 1
	For Each item_ In nodelist

		bln = true
		if instr(processed_parents, "||" & getFieldValue(item_,"key") & "||") > 0 then
			bln = false
		end if

		if bln = true then
			xml = xml & "<item>" & vbcrlf
			For Each field_in_item_ In item_.ChildNodes
				xml = xml & "<" & field_in_item_.BaseName & ">" & xmlsafe("" & field_in_item_.Text) & "</" & field_in_item_.BaseName & ">" & vbcrlf
			Next
			xml = xml & "</item>" & vbcrlf
		end if

		i = i + 1
	Next
	xml = xml & "</item_result>" & vbcrlf
	objXML.LoadXML xml

	'ok now we have a good set of xml to go with
	i = 0
	response.write "Items below do not have an action defined yet<br>"
	response.write "<table style='font-family:verdana;font-size:8pt;border-collapse:collapse;' border='1' id='table_2'>"
	response.write "<tr>"
	response.write "<td>" & "Parent Action id" & "</td>"
	response.write "<td>" & "Parent Action issuetype" & "</td>"
	response.write "<td>" & "Parent Action summary" & "</td>"
	response.write "<td>" & "Parent Action assignee" & "</td>"
	response.write "<td>" & "Parent Action created" & "</td>"
	response.write "</tr>"
	Set nodelist = objXML.getElementsByTagName("item_result/*")
	For Each item_ In nodelist
		response.write "<td>" & getFieldValue(item_,"key") & "</td>"
		response.write "<td>" & getFieldValue(item_,"issuetype") & "</td>"
		response.write "<td>" & getFieldValue(item_,"summary") & "</td>"
		response.write "<td>" & getFieldValue(item_,"assignee") & "</td>"
		created = getFieldValue(item_,"created")
		response.write "<td>" & toDD_MM_YYYY(created) & "</td>"
		response.write "</tr>"
		i = i + 1
	next

	response.write "</table>"
	response.write "<a href='#' onclick='downloadCSV(""table_2"",""output.csv"");'>CSV</a><br>"
	response.write i & " items<br>" ' were set to Closed in " & yearmonth & " however " & j & " items have Quality Declined<br>"
	response.write "<br>"

%>
</body>
</html>
<%
function showFirst(s) '20181227||11/Jan/19||2/Jan/19@@20181227||2/Jan/19||11/Jan/19 comes in
showFirst = ""
dim arr
dim arr2
dim a
dim tmp
if s = "" then
else
	arr = split(s, "@@")
	arr2 = split(arr(0), "||")
	showFirst = arr2(1)
end if
end function

function showLast(s) '20181227||11/Jan/19||2/Jan/19@@20181227||2/Jan/19||11/Jan/19 comes in
showLast = ""
dim arr
dim arr2
dim a
dim tmp
if s = "" then
else
	arr = split(s, "@@")
	arr2 = split(arr(ubound(arr)-1), "||")
	showLast = arr2(2)
end if
end function

function countChanges(s) '20181227||11/Jan/19||2/Jan/19@@20181227||2/Jan/19||11/Jan/19 comes in
countChanges = ""
dim arr
dim arr2
dim a
dim tmp
if s = "" then
else
	arr = split(s, "@@")
	countChanges = ubound(arr)
end if
end function

function mmmTOmm(s)
mmmTOmm = s
mmmTOmm = replace(mmmTOmm, "/Jan/", "01")
mmmTOmm = replace(mmmTOmm, "/Feb/", "02")
mmmTOmm = replace(mmmTOmm, "/Mar/", "03")
mmmTOmm = replace(mmmTOmm, "/Apr/", "04")
mmmTOmm = replace(mmmTOmm, "/May/", "05")
mmmTOmm = replace(mmmTOmm, "/Jun/", "06")
mmmTOmm = replace(mmmTOmm, "/Jul/", "07")
mmmTOmm = replace(mmmTOmm, "/Aug/", "08")
mmmTOmm = replace(mmmTOmm, "/Sep/", "09")
mmmTOmm = replace(mmmTOmm, "/Oct/", "10")
mmmTOmm = replace(mmmTOmm, "/Nov/", "11")
mmmTOmm = replace(mmmTOmm, "/Dec/", "12")
end function

function countDaysbetween(s) '20181227||11/Jan/19||2/Jan/19@@20181227||2/Jan/19||11/Jan/19 comes in
'update 14APR2020 : dateformat changed what comes in : yyyy-mm-dd
countDaysbetween = ""
Dim dt1
Dim dt2
dt1 = showFirst(s)
dt2 = showLast(s)
If dt1 = "" Or dt2 = "" Then
Else
	'dt1 = right("000000" & mmmTOmm(dt1),6)
	'dt2 = right("000000" & mmmTOmm(dt2),6)
	'countDaysbetween = dt1 & "#" & dt2 & "#"
	'dt1 = "20" & Mid(dt1, 5, 2) & Mid(dt1, 3, 2) & Mid(dt1, 1, 2)
	'dt2 = "20" & Mid(dt2, 5, 2) & Mid(dt2, 3, 2) & Mid(dt2, 1, 2)
	'14APR - just replace the - with space and that 'll give YYYYMMDD
	dt1 = replace(dt1, "-", "")
	dt2 = replace(dt2, "-", "")

	'countDaysbetween = countDaysbetween & dt1 & "#" & dt2 & "#"
	countDaysbetween = DateDiff("d", toDate(dt1), toDate(dt2))
End If
end function

function toYYYYMMDD(s)
toYYYYMMDD = s
if toYYYYMMDD <> "" then
	toYYYYMMDD = mid(s,1,4) & mid(s,6,2) & mid(s,9,2)
end if
end function

Function Lng2Dt(n)
if n = "" then
Lng2Dt = ""
else
n=n/1000
Lng2Dt = dateserial(1900,1,1) + (n + 2 ^ 31) / 86400
'Lng2Dt = right("0000" & year(lng2dt),4) & right("00" & month(lng2dt),2) & right("00" & day(lng2dt),2) & right("00" & hour(lng2dt),2) & right("00" & minute(lng2dt),2) & right("00" & second(lng2dt),2)
end if
End Function

function toDate(s)
'toDate = timeserial(mid(s,1,4),mid(s,5,2),mid(s,7,2),mid(s,9,2),mid(s,11,2),mid(s,13,2))
toDate = dateserial(mid(s,1,4),mid(s,5,2),mid(s,7,2))
end function

Function URLEncode(ByVal str)
 Dim strTemp, strChar
 Dim intPos, intASCII
 strTemp = ""
 strChar = ""
 For intPos = 1 To Len(str)
  intASCII = Asc(Mid(str, intPos, 1))
  If intASCII = 32 Then
   strTemp = strTemp & "+"
  ElseIf ((intASCII < 123) And (intASCII > 96)) Then
   strTemp = strTemp & Chr(intASCII)
  ElseIf ((intASCII < 91) And (intASCII > 64)) Then
   strTemp = strTemp & Chr(intASCII)
  ElseIf ((intASCII < 58) And (intASCII > 47)) Then
   strTemp = strTemp & Chr(intASCII)
  Else
   strChar = Trim(Hex(intASCII))
   If intASCII < 16 Then
    strTemp = strTemp & "%0" & strChar
   Else
    strTemp = strTemp & "%" & strChar
   End If
  End If
 Next
 URLEncode = strTemp
End Function

function getMostRecentDecision(all)
getMostRecentDecision = "99999999"
dim arr
dim arr2
dim a

arr = split(all, "@@")
for a = lbound(arr) to ubound(arr)
if arr(a) <> "" then
	arr2 = split(arr(a), "||")
	if arr2(0) < getMostRecentDecision then
		getMostRecentDecision = arr2(0)
	end if
end if
next
if getMostRecentDecision = "99999999" then getMostRecentDecision = ""
end function

function getSettoClosed(allstates)
getSettoClosed = "00000000"
dim arr
dim arr2
dim a

arr = split(allstates, "@@")
for a = lbound(arr) to ubound(arr)
if arr(a) <> "" then
	arr2 = split(arr(a), "||")
	if arr2(2) = "Closed" then
	getSettoClosed = arr2(0)
	end if
end if
next
if getSettoClosed = "00000000" then getSettoClosed = ""
end function

function donethismonth(allstates, ym)
donethismonth = false
'allstates = date||from||to@@...
dim arr
dim arr2
dim a

arr = split(allstates, "@@")
for a = lbound(arr) to ubound(arr)
if arr(a) <> "" then
	arr2 = split(arr(a), "||")
	'response.write ym & " - " & arr2(0) & "-" & arr2(1) & "-" & arr2(2) & "-" & "<br>"
	if arr2(2) = "Done" then
	if mid(arr2(0), 1, 6) = ym then
		donethismonth = true
	end if
	end if
end if
next
end function

function xmlsafe(s)
xmlsafe = s
xmlsafe = replace(xmlsafe, "&", "&amp;")
xmlsafe = replace(xmlsafe, "<", "&lt;")
xmlsafe = replace(xmlsafe, ">", "&gt;")
xmlsafe = replace(xmlsafe, chr(11), "")
end function

Function getJiraItems(jql)
randomize timer
'On Error Resume Next
Dim xmlhttp
dim url
url = "https://soreporting.azurewebsites.net/so_reporting/getJiraitems.aspx?jql=" & jql & "&rnd=" & rnd & ""
Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
xmlhttp.Open "GET", url, False
xmlhttp.send
getJiraItems = xmlhttp.responseText
Set xmlhttp = Nothing
End Function

function getfieldValue(xml, fld)
dim field_in_item_
For Each field_in_item_ In xml.ChildNodes
if field_in_item_.basename = fld then
getfieldValue = field_in_item_.text
end if
next
end function

function itemRejected(allstates)
'in = 20180226||Open||Breakdown@@20180226||Breakdown||Breakdown Done@@20180302||Breakdown Done||Development@@20180302||Development||Development Done@@20180305||Development Done||Validation@@20180305||Validation||Done
dim arr
dim arr2
dim a
itemRejected = false
arr = split(allstates, "@@")
for a = lbound(arr) to ubound(arr)
if arr(a) <> "" then
	arr2 = split(arr(a), "||")
	if arr2(1) = "Validation" and arr2(2) <> "Done" then
	itemRejected = true
	end if
end if
next
end function

function toDD_MM_YYYY(s)
toDD_MM_YYYY = s
if toDD_MM_YYYY = "" then exit function
'20180526 comes in
'26/05/2018 goes out
toDD_MM_YYYY = mid(s, 7, 2) & "/" & mid(s, 5, 2) &"/" & mid(s, 1, 4)
end function
%>