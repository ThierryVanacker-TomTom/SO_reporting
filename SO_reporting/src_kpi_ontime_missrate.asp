<%option explicit%>
<%server.scripttimeout=10*60 '10minutes%>
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

dim fso
dim f

arrP = split("STSPES,STSCNT,MOEBIUS,RES", ",")
'arr = split("STSPES", ",")
'arr = split("STSCNT,STSPES", ",")
arrP = split(request.querystring("project"), ",")
'response.write request.querystring("project") & "<br>"
'response.write ubound(arr) & "<br>"

dim dt
dim from_
dim to_

yearmonth = request.querystring("yearmonth")
if (yearmonth = "") then yearmonth = right("0000" & year(date),4) & right("00" & Month(date),2)

dt = dateserial(mid(yearmonth,1,4), mid(yearmonth,5,2), 1)
'format into 2018/06/30
from_ = right("0000" & year(dt),4) & "/" & right("00" & month(dt),2) & "/" & right("00" & day(dt),2)
dt = dateadd("m", 1, dt)
dt = dateadd("d", -1, dt)
to_ = right("0000" & year(dt),4) & "/" & right("00" & month(dt),2) & "/" & right("00" & day(dt),2)

dim jql
jql = ""
jql = jql & " project = ""OM"" "
' jql = jql & " and issuekey = ""OM-10666"" "
jql = jql & " AND issuetype in (""Activity - Archive Validation"", ""Activity - Feasibility"", ""Activity - Manual Production"", ""Activity - Other"", ""Activity - Quality Analysis"", ""Activity - SQR Measurement"", ""Activity - Source Acquisition"", ""Activity - Source Acquisition - Field"", ""Activity - Source Analysis"", ""Activity - Source Preparation"")"
jql = jql & " AND ""Assigned Unit"" in (""SO MOMA"",""SO SSO"",""SO APA"",""SO AME"",""SO PMO"",""SO GDT"",""SO EAP"", ""SO ECA"", ""SO SAMEA"",""SO EECA"", ""SO AFR"", ""SO WCE"", ""SO STS"", ""SO SAM"", ""SO PDV"", ""SO OCE"", ""SO NEA"", ""SO NAM"", ""SO LAM"")"
jql = jql & " AND ("
' ' jql = jql & " status changed to done during (""" & from_ & """, """ & to_ & """)"
jql = jql & " status changed to ""closed"" during (""" & from_ & """, """ & to_ & """)"
jql = jql & " OR status changed to ""done"" during (""" & from_ & """, """ & to_ & """)"
' 'jql = jql & " OR status changed to planned during (""" & from_ & """"", """ & to_ & """)"
jql = jql & " )"


' response.write jql
' response.end

xml = getJiraItems(jql)
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
'	response.write ("<br>" & replace(xml, "<" , "["))
	objXML.LoadXML xml

	'ok now we have a good set of xml to go with
	i = 0
	j = 0
	response.write "<style='font-family:verdana;font-size:12pt;><b>OTD Miss Rate</b>"
	response.write "<br><style='font-family:verdana;font-size:10pt;>All issues of type (Activity - Archive Validation,  Activity - Feasibility,  Activity - Manual Production,  Activity - Other,  Activity - Quality Analysis,  Activity - SQR Measurement,  Activity - Source Acquisition,  Activity - Source Acquisition - Field,  Activity - Source Analysis,  Activity - Source Preparation) "
	response.write"<br><style='font-family:verdana;font-size:10pt;> and Assigned Unit of (SO MOMA, SO SSO, SO APA, SO AME, SO PMO, SO GDT, SO EAP, SO ECA, SO SAMEA, SO EECA,  SO AFR,  SO WCE,  SO STS,  SO SAM,  SO PDV,  SO OCE,  SO NEA,  SO NAM,  SO LAM) "
	response.write"<br><style='font-family:verdana;font-size:10pt;>closed in month XX. Review for missed End Date."
	response.write "<table style='font-family:verdana;font-size:8pt;border-collapse:collapse;' border='1' id='table'>"
	response.write "<tr>"
	response.write "<td>" & "pkey" & "-" & "issuenum" & "</td>"
	response.write "<td>" & "issuetype" & "</td>"
	' response.write "<td>" & "component" & "</td>"
	response.write "<td>" & "summary" & "</td>"
	response.write "<td>" & "state"   & "</td>"
	response.write "<td>" & "assigned unit"   & "</td>"
	response.write "<td>" & "requesting unit"   & "</td>"
	response.write "<td>" & "Baseline end date"   & "</td>"
	response.write "<td>" & "End date"   & "</td>"
	' response.write "<td>" & "resolution"   & "</td>"
	' response.write "<td>" & "resolutiondate"   & "</td>"
	response.write "<td>" & "state change overview"   & "</td>"
	response.write "<td>" & "date set to Done"   & "</td>"
	response.write "<td>" & "date set to closed"   & "</td>"
	response.write "<td>" & "created"   & "</td>"
	response.write "<td>" & "updated"   & "</td>"
	' response.write "<td>" & "item duedate"   & "</td>"
	' response.write "<td>" & "item target duedate"   & "</td>"
	' response.write "<td>" & "duedate via linked sprint(s)"   & "</td>"
	response.write "</tr>"
	Set nodelist = objXML.getElementsByTagName("item_result/*")
	For Each item_ In nodelist
		bln = true
		duedate = getFieldValue(item_,"enddate")
		target_duedate = getFieldValue(item_,"bl_enddate")

		if bln = true then
		if duedate > target_duedate then
			j = j + 1
			response.write "<tr style='background-color:red'>"
			else
			response.write "<tr>"
			end if
			response.write "<td>" & getFieldValue(item_,"key") & "</td>"
			response.write "<td>" & getFieldValue(item_,"issuetype") & "</td>"
			' response.write "<td>" & getFieldValue(item_,"components") & "</td>"
			response.write "<td>" & getFieldValue(item_,"summary") & "</td>"
			response.write "<td>" & getFieldValue(item_,"status") & "</td>"
			response.write "<td>" & getFieldValue(item_,"assigned_unit") & "</td>"
			response.write "<td>" & getFieldValue(item_,"requesting_unit") & "</td>"
			response.write "<td>" & toDD_MM_YYYY(getFieldValue(item_,"bl_enddate")) & "</td>"
			response.write "<td>" & toDD_MM_YYYY(getFieldValue(item_,"enddate")) & "</td>"
			'response.write "<td>" & getFieldValue(item_,"components") & "</td>"
			' response.write "<td>" & getFieldValue(item_,"resolution") & "</td>"
			' response.write "<td>" & toDD_MM_YYYY(getFieldValue(item_,"resolutiondate")) & "</td>"
			response.write "<td>" & replace(getFieldValue(item_,"transitions_history"), "@@", "<br>") & "</td>"
			response.write "<td>" & toDD_MM_YYYY(getSettoDone(getFieldValue(item_,"transitions_history"))) & "</td>"
			response.write "<td>" & toDD_MM_YYYY(getSettoClosed(getFieldValue(item_,"transitions_history"))) & "</td>"
			response.write "<td>" & toDD_MM_YYYY(getFieldValue(item_,"created")) & "</td>"
			response.write "<td>" & toDD_MM_YYYY(getFieldValue(item_,"updated")) & "</td>"
			'response.write "<td>" & "" & "</td>" 'validation tracking to in validation occurences
			'response.write "<td>" & "" & "</td>" 'in validation to anything but done/Validation Tracking occurences
			' response.write "<td>" & toDD_MM_YYYY(getFieldValue(item_,"duedate")) & "</td>"
			' response.write "<td>" & toDD_MM_YYYY(getFieldValue(item_,"target_duedate")) & "</td>"
			' response.write "<td>" & toDD_MM_YYYY(getFieldValue(item_,"sprintcompletedate")) & "</td>"

			'For Each field_in_item_ In item_.ChildNodes
			'response.write "<td>" & field_in_item_.text & "</td>"
			'next
			response.write "</tr>"
			i = i + 1
		end if
	next

	response.write "</table>"
	response.write "<a href='#' onclick='downloadCSV(""table"",""output.csv"");'>CSV</a><br>"
	response.write i & " items were set to closed in " & yearmonth & " however " & j & " items have been rejected<br>"
	'response.write "<br>"

'end if 'prj loop
'next 'prj loop

'response.write tmp
%>
</body>
</html>
<%
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

function getSettoDone(allstates)
getSettoDone = "00000000"
dim arr
dim arr2
dim a

arr = split(allstates, "@@")
for a = lbound(arr) to ubound(arr)
if arr(a) <> "" then
	arr2 = split(arr(a), "||")
	if arr2(2) = "Done" then
	getSettoDone = arr2(0)
	end if
end if
next
if getSettoDone = "00000000" then getSettoDone = ""
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
'response.write "http://127.0.0.1/src_reporting/getJiraitems.aspx?project=" & project & "&yearmonth=" & yearmonth & "&rnd=" & rnd & ""
randomize timer
'On Error Resume Next
Dim xmlhttp
Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
' response.write  "http://127.0.0.1/amy/src_reporting/getJiraitems.aspx?jql=" & jql & "&rnd=" & rnd & ""
' response.end
xmlhttp.Open "GET", "http://127.0.0.1/so_reporting/getJiraitems.aspx?jql=" & jql & "&rnd=" & rnd & "", False
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