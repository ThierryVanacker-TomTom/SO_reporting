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
'jql = jql & " and issuekey in (OM-11393, OM-11348, OM-11334, OM-10670, OM-10667, OM-10666, OM-10665, OM-10664, OM-10663, OM-10648, OM-10623, OM-10607, OM-7007, OM-6999, OM-6991, OM-6990, OM-6989, OM-6986, OM-6983, OM-6980, OM-6966, OM-6961, OM-6960, OM-6959, OM-6956, OM-6955, OM-6952, OM-6951, OM-6950, OM-6946, OM-6945, OM-6944, OM-6941, OM-6940, OM-6938, OM-6936, OM-6930, OM-6929, OM-6928, OM-6927, OM-6926, OM-6913, OM-6906, OM-6905, OM-6903, OM-6902, OM-6901, OM-6900, OM-6899, OM-6898, OM-6890, OM-6883, OM-6878, OM-6874, OM-6824, OM-6819, OM-6817, OM-6813, OM-6804, OM-6803, OM-6800, OM-6797, OM-6790, OM-6789, OM-6756, OM-6751, OM-6747, OM-6737, OM-6733, OM-6730, OM-6725, OM-6723, OM-6713, OM-6711, OM-6709, OM-6708, OM-6706, OM-6705, OM-6703, OM-6702, OM-6696, OM-6694, OM-6693, OM-6692, OM-6689, OM-6688, OM-6685, OM-6684, OM-6683, OM-6682, OM-6679, OM-6676, OM-6675, OM-6674, OM-6670, OM-6665, OM-6664, OM-6663, OM-6662, OM-6661, OM-6660, OM-6659, OM-6653, OM-6652, OM-6651, OM-6644, OM-6640, OM-6635, OM-6631, OM-6627, OM-6622, OM-6620, OM-6618, OM-6612, OM-6599, OM-6593, OM-6590, OM-6585, OM-6584, OM-6580, OM-6579, OM-6577, OM-6573, OM-6572, OM-6571, OM-6570, OM-6569, OM-6567, OM-6566, OM-6561, OM-6556, OM-6552, OM-6549, OM-6547, OM-6546, OM-6543, OM-6539, OM-6532, OM-6531, OM-6529, OM-6527, OM-6525, OM-6524, OM-6522, OM-6520, OM-6514, OM-6511, OM-6509, OM-6508, OM-6506, OM-6505, OM-6501, OM-6500, OM-6499, OM-6498, OM-6497, OM-6495, OM-6494, OM-6493, OM-6491, OM-6490, OM-6489, OM-6488, OM-6487, OM-6466, OM-6465, OM-6464, OM-6462, OM-6442, OM-6435, OM-6434, OM-6433, OM-6431, OM-6430, OM-6429, OM-6428, OM-6427, OM-6426, OM-6425, OM-6424, OM-6423, OM-6422, OM-6421, OM-6419, OM-6418, OM-6417, OM-6416, OM-6414, OM-6405, OM-6397, OM-6393, OM-6391, OM-6389, OM-6385, OM-6383, OM-6382, OM-6381, OM-6376, OM-6374, OM-6373, OM-6372, OM-6371, OM-6370, OM-6369, OM-6368, OM-6366, OM-6365, OM-6364, OM-6363, OM-6347, OM-6331, OM-6330, OM-6329, OM-6328, OM-6327, OM-6326, OM-6324, OM-6323, OM-6322, OM-6321, OM-6319, OM-6318, OM-6316, OM-6314, OM-6313, OM-6311, OM-6310, OM-6309, OM-6308, OM-6307, OM-6306, OM-6305, OM-6304, OM-6303, OM-6302, OM-6301, OM-6300, OM-6296, OM-6295, OM-6294, OM-6293, OM-6292, OM-6291, OM-6288, OM-6285, OM-6284, OM-6282, OM-6281, OM-6280, OM-6279, OM-6271, OM-6269, OM-6268, OM-6267, OM-6266, OM-6263, OM-6255, OM-6254, OM-6253, OM-6247, OM-6241, OM-6240, OM-6239, OM-6237, OM-6234, OM-6232, OM-6231, OM-6230, OM-6229, OM-6227, OM-6224, OM-6223, OM-6222, OM-6221, OM-6220, OM-6215, OM-6214, OM-6213, OM-6212, OM-6211, OM-6210, OM-6209, OM-6207, OM-6206, OM-6205, OM-6204, OM-6203, OM-6200, OM-6199, OM-6198, OM-6197, OM-6196, OM-6193, OM-6187, OM-6184, OM-6183, OM-6182, OM-6177, OM-6172, OM-6170, OM-6166, OM-6163, OM-6162, OM-6161, OM-6159, OM-6155, OM-6143, OM-6142, OM-6141, OM-6140, OM-6139, OM-6138, OM-6137, OM-6134, OM-6132, OM-6130, OM-6124, OM-6121, OM-6120, OM-6119, OM-6115, OM-6109, OM-6098, OM-6097, OM-6093, OM-6092, OM-6091, OM-6090, OM-6079, OM-6075, OM-6074, OM-6073, OM-6065, OM-6054, OM-6044, OM-6043, OM-6037, OM-6027, OM-6023, OM-6020, OM-6018, OM-6015, OM-6014, OM-6013, OM-6010, OM-6009, OM-6008, OM-6002, OM-6001, OM-5997, OM-5989, OM-5988, OM-5984, OM-5983, OM-5982, OM-5981, OM-5976, OM-5975, OM-5974, OM-5973, OM-5972, OM-5971, OM-5967, OM-5966, OM-5964, OM-5963, OM-5962, OM-5959, OM-5957, OM-5955, OM-5953, OM-5952, OM-5951, OM-5950, OM-5949, OM-5948, OM-5946, OM-5945, OM-5941, OM-5937, OM-5936, OM-5928, OM-5927, OM-5914, OM-5913, OM-5905, OM-5904, OM-5903, OM-5901, OM-5893, OM-5889, OM-5888, OM-5887, OM-5886, OM-5885, OM-5884, OM-5879, OM-5878, OM-5877, OM-5872, OM-5870, OM-5867, OM-5866, OM-5858, OM-5834, OM-5828, OM-5819, OM-5816, OM-5815, OM-5814, OM-5813, OM-5811, OM-5808, OM-5807, OM-5806, OM-5803, OM-5800, OM-5799, OM-5798, OM-5797, OM-5795, OM-5793, OM-5792, OM-5791, OM-5790, OM-5789, OM-5788, OM-5787, OM-5784, OM-5775, OM-5773, OM-5772, OM-5765, OM-5764, OM-5763, OM-5761, OM-5758, OM-5757, OM-5755, OM-5754, OM-5750, OM-5749, OM-5748, OM-5746, OM-5745, OM-5743, OM-5742, OM-5741, OM-5739, OM-5723, OM-5722, OM-5721, OM-5718, OM-5705, OM-5679, OM-5677, OM-5672, OM-5668, OM-5665, OM-5663, OM-5662, OM-5661, OM-5658, OM-5654, OM-5652, OM-5647, OM-5642, OM-5640, OM-5638, OM-5633, OM-5629, OM-5627, OM-5626, OM-5625, OM-5624, OM-5623, OM-5620, OM-5619, OM-5618, OM-5584, OM-5582, OM-5578, OM-5576, OM-5572, OM-5571, OM-5570, OM-5568, OM-5567, OM-5566, OM-5564, OM-5563, OM-5562, OM-5561, OM-5560, OM-5559, OM-5558, OM-5557, OM-5554, OM-5552, OM-5550, OM-5549, OM-5542, OM-5540, OM-5539, OM-5538, OM-5537, OM-5536, OM-5526, OM-5520, OM-5519, OM-5518, OM-5516, OM-5515, OM-5514, OM-5513, OM-5512, OM-5510, OM-5508, OM-5507, OM-5505, OM-5500, OM-5499, OM-5498, OM-5493, OM-5492, OM-5486, OM-5482, OM-5456, OM-5432, OM-5429, OM-5427, OM-5426, OM-5422, OM-5398, OM-5389, OM-5386, OM-5384, OM-5377, OM-5376, OM-5375, OM-5374, OM-5373, OM-5372, OM-5349, OM-5348, OM-5346, OM-5344, OM-5342, OM-5341, OM-5340, OM-5339, OM-5337, OM-5336, OM-5335, OM-5334, OM-5330, OM-5327, OM-5312, OM-5285, OM-5271, OM-5270, OM-5269, OM-5268, OM-5267, OM-5266, OM-5265, OM-5264, OM-5257, OM-5254, OM-5251, OM-5249, OM-5246, OM-5245, OM-5244, OM-5243, OM-5242, OM-5238, OM-5237, OM-5236, OM-5235, OM-5230, OM-5228, OM-5222, OM-5221, OM-5210, OM-5206, OM-5202, OM-5200, OM-5196, OM-5192, OM-5191, OM-5190, OM-5185, OM-5182, OM-5181, OM-5180, OM-5179, OM-5178, OM-5177, OM-5174, OM-5173, OM-5171, OM-5170, OM-5169, OM-5168, OM-5167, OM-5165, OM-5164, OM-5163, OM-5162, OM-5161, OM-5160, OM-5155, OM-5133, OM-5131, OM-5130, OM-5129, OM-5128, OM-5127, OM-5126, OM-5125, OM-5124, OM-5118, OM-5117, OM-5114, OM-5113, OM-5112, OM-5111, OM-5110, OM-5109, OM-5089, OM-5088, OM-5087, OM-5072, OM-5071, OM-5070, OM-5069, OM-5066, OM-5062, OM-5061, OM-5060, OM-5058, OM-5057, OM-5055, OM-5053, OM-5051, OM-5050, OM-5044, OM-5043, OM-5041, OM-5040, OM-5028, OM-5027, OM-5023, OM-5021, OM-5020, OM-5017, OM-5016, OM-5015, OM-5009, OM-5004, OM-5003, OM-4999, OM-4997, OM-4996, OM-4994, OM-4992, OM-4991, OM-4989, OM-4988, OM-4987, OM-4985, OM-4984, OM-4982, OM-4981, OM-4980, OM-4979, OM-4978, OM-4977, OM-4976, OM-4975, OM-4974, OM-4973, OM-4972, OM-4970, OM-4967, OM-4966, OM-4965, OM-4964, OM-4959, OM-4954, OM-4951, OM-4950, OM-4949, OM-4947, OM-4946, OM-4945, OM-4944, OM-4943, OM-4942, OM-4941, OM-4940, OM-4939, OM-4936, OM-4933, OM-4931, OM-4930, OM-4929, OM-4927, OM-4923, OM-4921, OM-4920, OM-4917, OM-4913, OM-4912, OM-4910, OM-4909, OM-4908, OM-4907, OM-4906, OM-4905, OM-4904, OM-4903, OM-4901, OM-4900, OM-4898, OM-4897, OM-4896, OM-4895, OM-4894, OM-4893, OM-4892, OM-4891, OM-4890, OM-4889, OM-4887, OM-4886, OM-4885, OM-4882, OM-4880, OM-4879, OM-4877, OM-4875, OM-4873, OM-4872, OM-4871, OM-4870, OM-4869, OM-4867, OM-4866, OM-4865, OM-4864, OM-4856, OM-4855, OM-4854, OM-4853, OM-4852, OM-4851, OM-4850, OM-4849, OM-4848, OM-4847, OM-4846, OM-4845, OM-4835, OM-4833, OM-4832, OM-4827, OM-4826, OM-4823, OM-4822, OM-4821, OM-4819, OM-4818, OM-4817, OM-4816, OM-4815, OM-4814, OM-4813, OM-4812, OM-4811, OM-4810, OM-4809, OM-4808, OM-4807, OM-4806, OM-4805, OM-4801, OM-4800, OM-4799, OM-4798, OM-4797, OM-4796, OM-4795, OM-4794, OM-4793, OM-4792, OM-4791, OM-1000, OM-995, OM-989, OM-984, OM-981, OM-980, OM-979, OM-978, OM-977, OM-976, OM-973, OM-971, OM-970, OM-964, OM-963, OM-962, OM-961, OM-958, OM-955, OM-947, OM-946, OM-945, OM-941, OM-940, OM-939, OM-938, OM-936, OM-935, OM-934, OM-933, OM-932, OM-928, OM-927, OM-926, OM-925, OM-924, OM-921, OM-920, OM-919, OM-913, OM-910, OM-908, OM-890, OM-889, OM-883, OM-881, OM-880, OM-879, OM-878, OM-877, OM-876, OM-875, OM-874, OM-871, OM-870, OM-869, OM-868, OM-866, OM-865, OM-863, OM-861, OM-860, OM-859, OM-858, OM-857, OM-856, OM-855, OM-853, OM-852, OM-850, OM-848, OM-847, OM-846, OM-843, OM-839, OM-836, OM-833, OM-829, OM-828, OM-827, OM-826, OM-825, OM-822, OM-820, OM-819, OM-818, OM-817, OM-815, OM-812, OM-811, OM-807, OM-804, OM-792, OM-784, OM-783, OM-781, OM-779, OM-775, OM-773, OM-768, OM-763, OM-762, OM-761, OM-760, OM-754, OM-753, OM-752, OM-750, OM-746, OM-742, OM-741, OM-740, OM-736, OM-735, OM-734, OM-723, OM-722, OM-719, OM-717, OM-715, OM-713, OM-710, OM-707, OM-705, OM-704, OM-702, OM-701, OM-699, OM-698, OM-697, OM-695, OM-694, OM-690, OM-688, OM-687, OM-686, OM-684, OM-683, OM-680, OM-679, OM-678, OM-677, OM-676, OM-670, OM-668, OM-665, OM-655, OM-654, OM-652, OM-651, OM-649, OM-647, OM-645, OM-644, OM-643, OM-642, OM-639, OM-638, OM-637, OM-636, OM-635, OM-634, OM-625) "
jql = jql & " AND issuetype in (""Activity - Archive Validation"", ""Activity - Feasibility"", ""Activity - Manual Production"", ""Activity - Other"", ""Activity - Quality Analysis"", ""Activity - SQR Measurement"", ""Activity - Source Acquisition"", ""Activity - Source Acquisition - Field"", ""Activity - Source Analysis"", ""Activity - Source Preparation"")"
jql = jql & " AND ""Assigned Unit"" in (""SO MOMA"",""SO SSO"",""SO APA"",""SO AME"",""SO PMO"",""SO GDT"",""SO EAP"", ""SO ECA"", ""SO SAMEA"",""SO EECA"", ""SO AFR"", ""SO WCE"", ""SO STS"", ""SO SAM"", ""SO PDV"", ""SO OCE"", ""SO NEA"", ""SO NAM"", ""SO LAM"")"
jql = jql & " AND ("
' ' jql = jql & " status changed to done during (""" & from_ & """, """ & to_ & """)"
jql = jql & " status changed to done during (""" & from_ & """, """ & to_ & """)"
' jql = jql & " OR status changed to ""done"" during (""" & from_ & """, """ & to_ & """)"
' 'jql = jql & " OR status changed to planned during (""" & from_ & """"", """ & to_ & """)"
jql = jql & " )"
jql = jql & " and ("
jql = jql & " resolution in (""Accepted"",""Unresolved"",""Rejected"") "
jql = jql & " ) "
dim expl
expl = ""
expl = expl & "<span style='font-family:verdana;font-size:10pt;'><b>OTD Miss Rate</b></span><br>"
expl = expl & "<span style='font-family:verdana;font-size:8pt;'>"
expl = expl & "All issues of type (Activity - Archive Validation,  Activity - Feasibility,  Activity - Manual Production,  Activity - Other,  Activity - Quality Analysis,  Activity - SQR Measurement,  Activity - Source Acquisition,  Activity - Source Acquisition - Field,  Activity - Source Analysis,  Activity - Source Preparation)<br>"
expl = expl & "and Assigned Unit of (SO MOMA, SO SSO, SO APA, SO AME, SO PMO, SO GDT, SO EAP, SO ECA, SO SAMEA,SO EECA,  SO AFR,  SO WCE,  SO STS,  SO SAM,  SO PDV,  SO OCE,  SO NEA,  SO NAM,  SO LAM)<br>"
expl = expl & "and set to done in month XX<br>"
expl = expl & "and resolution in Accepted/Rejected/Unresolved.<br>"
expl = expl & "Review for missed End Date.<br>"
expl = expl & "</span>"

'if we run for 2019 month we need to use a different query (also use a different explanantion for the top)
if mid(yearmonth, 1, 4) <> "2018" then
	jql = ""
	jql = jql & " project = ""OM"" "
	' jql = jql & " and issuekey = ""OM-10666"" "
	'jql = jql & " and issuekey in (OM-11393, OM-11348, OM-11334, OM-10670, OM-10667, OM-10666, OM-10665, OM-10664, OM-10663, OM-10648, OM-10623, OM-10607, OM-7007, OM-6999, OM-6991, OM-6990, OM-6989, OM-6986, OM-6983, OM-6980, OM-6966, OM-6961, OM-6960, OM-6959, OM-6956, OM-6955, OM-6952, OM-6951, OM-6950, OM-6946, OM-6945, OM-6944, OM-6941, OM-6940, OM-6938, OM-6936, OM-6930, OM-6929, OM-6928, OM-6927, OM-6926, OM-6913, OM-6906, OM-6905, OM-6903, OM-6902, OM-6901, OM-6900, OM-6899, OM-6898, OM-6890, OM-6883, OM-6878, OM-6874, OM-6824, OM-6819, OM-6817, OM-6813, OM-6804, OM-6803, OM-6800, OM-6797, OM-6790, OM-6789, OM-6756, OM-6751, OM-6747, OM-6737, OM-6733, OM-6730, OM-6725, OM-6723, OM-6713, OM-6711, OM-6709, OM-6708, OM-6706, OM-6705, OM-6703, OM-6702, OM-6696, OM-6694, OM-6693, OM-6692, OM-6689, OM-6688, OM-6685, OM-6684, OM-6683, OM-6682, OM-6679, OM-6676, OM-6675, OM-6674, OM-6670, OM-6665, OM-6664, OM-6663, OM-6662, OM-6661, OM-6660, OM-6659, OM-6653, OM-6652, OM-6651, OM-6644, OM-6640, OM-6635, OM-6631, OM-6627, OM-6622, OM-6620, OM-6618, OM-6612, OM-6599, OM-6593, OM-6590, OM-6585, OM-6584, OM-6580, OM-6579, OM-6577, OM-6573, OM-6572, OM-6571, OM-6570, OM-6569, OM-6567, OM-6566, OM-6561, OM-6556, OM-6552, OM-6549, OM-6547, OM-6546, OM-6543, OM-6539, OM-6532, OM-6531, OM-6529, OM-6527, OM-6525, OM-6524, OM-6522, OM-6520, OM-6514, OM-6511, OM-6509, OM-6508, OM-6506, OM-6505, OM-6501, OM-6500, OM-6499, OM-6498, OM-6497, OM-6495, OM-6494, OM-6493, OM-6491, OM-6490, OM-6489, OM-6488, OM-6487, OM-6466, OM-6465, OM-6464, OM-6462, OM-6442, OM-6435, OM-6434, OM-6433, OM-6431, OM-6430, OM-6429, OM-6428, OM-6427, OM-6426, OM-6425, OM-6424, OM-6423, OM-6422, OM-6421, OM-6419, OM-6418, OM-6417, OM-6416, OM-6414, OM-6405, OM-6397, OM-6393, OM-6391, OM-6389, OM-6385, OM-6383, OM-6382, OM-6381, OM-6376, OM-6374, OM-6373, OM-6372, OM-6371, OM-6370, OM-6369, OM-6368, OM-6366, OM-6365, OM-6364, OM-6363, OM-6347, OM-6331, OM-6330, OM-6329, OM-6328, OM-6327, OM-6326, OM-6324, OM-6323, OM-6322, OM-6321, OM-6319, OM-6318, OM-6316, OM-6314, OM-6313, OM-6311, OM-6310, OM-6309, OM-6308, OM-6307, OM-6306, OM-6305, OM-6304, OM-6303, OM-6302, OM-6301, OM-6300, OM-6296, OM-6295, OM-6294, OM-6293, OM-6292, OM-6291, OM-6288, OM-6285, OM-6284, OM-6282, OM-6281, OM-6280, OM-6279, OM-6271, OM-6269, OM-6268, OM-6267, OM-6266, OM-6263, OM-6255, OM-6254, OM-6253, OM-6247, OM-6241, OM-6240, OM-6239, OM-6237, OM-6234, OM-6232, OM-6231, OM-6230, OM-6229, OM-6227, OM-6224, OM-6223, OM-6222, OM-6221, OM-6220, OM-6215, OM-6214, OM-6213, OM-6212, OM-6211, OM-6210, OM-6209, OM-6207, OM-6206, OM-6205, OM-6204, OM-6203, OM-6200, OM-6199, OM-6198, OM-6197, OM-6196, OM-6193, OM-6187, OM-6184, OM-6183, OM-6182, OM-6177, OM-6172, OM-6170, OM-6166, OM-6163, OM-6162, OM-6161, OM-6159, OM-6155, OM-6143, OM-6142, OM-6141, OM-6140, OM-6139, OM-6138, OM-6137, OM-6134, OM-6132, OM-6130, OM-6124, OM-6121, OM-6120, OM-6119, OM-6115, OM-6109, OM-6098, OM-6097, OM-6093, OM-6092, OM-6091, OM-6090, OM-6079, OM-6075, OM-6074, OM-6073, OM-6065, OM-6054, OM-6044, OM-6043, OM-6037, OM-6027, OM-6023, OM-6020, OM-6018, OM-6015, OM-6014, OM-6013, OM-6010, OM-6009, OM-6008, OM-6002, OM-6001, OM-5997, OM-5989, OM-5988, OM-5984, OM-5983, OM-5982, OM-5981, OM-5976, OM-5975, OM-5974, OM-5973, OM-5972, OM-5971, OM-5967, OM-5966, OM-5964, OM-5963, OM-5962, OM-5959, OM-5957, OM-5955, OM-5953, OM-5952, OM-5951, OM-5950, OM-5949, OM-5948, OM-5946, OM-5945, OM-5941, OM-5937, OM-5936, OM-5928, OM-5927, OM-5914, OM-5913, OM-5905, OM-5904, OM-5903, OM-5901, OM-5893, OM-5889, OM-5888, OM-5887, OM-5886, OM-5885, OM-5884, OM-5879, OM-5878, OM-5877, OM-5872, OM-5870, OM-5867, OM-5866, OM-5858, OM-5834, OM-5828, OM-5819, OM-5816, OM-5815, OM-5814, OM-5813, OM-5811, OM-5808, OM-5807, OM-5806, OM-5803, OM-5800, OM-5799, OM-5798, OM-5797, OM-5795, OM-5793, OM-5792, OM-5791, OM-5790, OM-5789, OM-5788, OM-5787, OM-5784, OM-5775, OM-5773, OM-5772, OM-5765, OM-5764, OM-5763, OM-5761, OM-5758, OM-5757, OM-5755, OM-5754, OM-5750, OM-5749, OM-5748, OM-5746, OM-5745, OM-5743, OM-5742, OM-5741, OM-5739, OM-5723, OM-5722, OM-5721, OM-5718, OM-5705, OM-5679, OM-5677, OM-5672, OM-5668, OM-5665, OM-5663, OM-5662, OM-5661, OM-5658, OM-5654, OM-5652, OM-5647, OM-5642, OM-5640, OM-5638, OM-5633, OM-5629, OM-5627, OM-5626, OM-5625, OM-5624, OM-5623, OM-5620, OM-5619, OM-5618, OM-5584, OM-5582, OM-5578, OM-5576, OM-5572, OM-5571, OM-5570, OM-5568, OM-5567, OM-5566, OM-5564, OM-5563, OM-5562, OM-5561, OM-5560, OM-5559, OM-5558, OM-5557, OM-5554, OM-5552, OM-5550, OM-5549, OM-5542, OM-5540, OM-5539, OM-5538, OM-5537, OM-5536, OM-5526, OM-5520, OM-5519, OM-5518, OM-5516, OM-5515, OM-5514, OM-5513, OM-5512, OM-5510, OM-5508, OM-5507, OM-5505, OM-5500, OM-5499, OM-5498, OM-5493, OM-5492, OM-5486, OM-5482, OM-5456, OM-5432, OM-5429, OM-5427, OM-5426, OM-5422, OM-5398, OM-5389, OM-5386, OM-5384, OM-5377, OM-5376, OM-5375, OM-5374, OM-5373, OM-5372, OM-5349, OM-5348, OM-5346, OM-5344, OM-5342, OM-5341, OM-5340, OM-5339, OM-5337, OM-5336, OM-5335, OM-5334, OM-5330, OM-5327, OM-5312, OM-5285, OM-5271, OM-5270, OM-5269, OM-5268, OM-5267, OM-5266, OM-5265, OM-5264, OM-5257, OM-5254, OM-5251, OM-5249, OM-5246, OM-5245, OM-5244, OM-5243, OM-5242, OM-5238, OM-5237, OM-5236, OM-5235, OM-5230, OM-5228, OM-5222, OM-5221, OM-5210, OM-5206, OM-5202, OM-5200, OM-5196, OM-5192, OM-5191, OM-5190, OM-5185, OM-5182, OM-5181, OM-5180, OM-5179, OM-5178, OM-5177, OM-5174, OM-5173, OM-5171, OM-5170, OM-5169, OM-5168, OM-5167, OM-5165, OM-5164, OM-5163, OM-5162, OM-5161, OM-5160, OM-5155, OM-5133, OM-5131, OM-5130, OM-5129, OM-5128, OM-5127, OM-5126, OM-5125, OM-5124, OM-5118, OM-5117, OM-5114, OM-5113, OM-5112, OM-5111, OM-5110, OM-5109, OM-5089, OM-5088, OM-5087, OM-5072, OM-5071, OM-5070, OM-5069, OM-5066, OM-5062, OM-5061, OM-5060, OM-5058, OM-5057, OM-5055, OM-5053, OM-5051, OM-5050, OM-5044, OM-5043, OM-5041, OM-5040, OM-5028, OM-5027, OM-5023, OM-5021, OM-5020, OM-5017, OM-5016, OM-5015, OM-5009, OM-5004, OM-5003, OM-4999, OM-4997, OM-4996, OM-4994, OM-4992, OM-4991, OM-4989, OM-4988, OM-4987, OM-4985, OM-4984, OM-4982, OM-4981, OM-4980, OM-4979, OM-4978, OM-4977, OM-4976, OM-4975, OM-4974, OM-4973, OM-4972, OM-4970, OM-4967, OM-4966, OM-4965, OM-4964, OM-4959, OM-4954, OM-4951, OM-4950, OM-4949, OM-4947, OM-4946, OM-4945, OM-4944, OM-4943, OM-4942, OM-4941, OM-4940, OM-4939, OM-4936, OM-4933, OM-4931, OM-4930, OM-4929, OM-4927, OM-4923, OM-4921, OM-4920, OM-4917, OM-4913, OM-4912, OM-4910, OM-4909, OM-4908, OM-4907, OM-4906, OM-4905, OM-4904, OM-4903, OM-4901, OM-4900, OM-4898, OM-4897, OM-4896, OM-4895, OM-4894, OM-4893, OM-4892, OM-4891, OM-4890, OM-4889, OM-4887, OM-4886, OM-4885, OM-4882, OM-4880, OM-4879, OM-4877, OM-4875, OM-4873, OM-4872, OM-4871, OM-4870, OM-4869, OM-4867, OM-4866, OM-4865, OM-4864, OM-4856, OM-4855, OM-4854, OM-4853, OM-4852, OM-4851, OM-4850, OM-4849, OM-4848, OM-4847, OM-4846, OM-4845, OM-4835, OM-4833, OM-4832, OM-4827, OM-4826, OM-4823, OM-4822, OM-4821, OM-4819, OM-4818, OM-4817, OM-4816, OM-4815, OM-4814, OM-4813, OM-4812, OM-4811, OM-4810, OM-4809, OM-4808, OM-4807, OM-4806, OM-4805, OM-4801, OM-4800, OM-4799, OM-4798, OM-4797, OM-4796, OM-4795, OM-4794, OM-4793, OM-4792, OM-4791, OM-1000, OM-995, OM-989, OM-984, OM-981, OM-980, OM-979, OM-978, OM-977, OM-976, OM-973, OM-971, OM-970, OM-964, OM-963, OM-962, OM-961, OM-958, OM-955, OM-947, OM-946, OM-945, OM-941, OM-940, OM-939, OM-938, OM-936, OM-935, OM-934, OM-933, OM-932, OM-928, OM-927, OM-926, OM-925, OM-924, OM-921, OM-920, OM-919, OM-913, OM-910, OM-908, OM-890, OM-889, OM-883, OM-881, OM-880, OM-879, OM-878, OM-877, OM-876, OM-875, OM-874, OM-871, OM-870, OM-869, OM-868, OM-866, OM-865, OM-863, OM-861, OM-860, OM-859, OM-858, OM-857, OM-856, OM-855, OM-853, OM-852, OM-850, OM-848, OM-847, OM-846, OM-843, OM-839, OM-836, OM-833, OM-829, OM-828, OM-827, OM-826, OM-825, OM-822, OM-820, OM-819, OM-818, OM-817, OM-815, OM-812, OM-811, OM-807, OM-804, OM-792, OM-784, OM-783, OM-781, OM-779, OM-775, OM-773, OM-768, OM-763, OM-762, OM-761, OM-760, OM-754, OM-753, OM-752, OM-750, OM-746, OM-742, OM-741, OM-740, OM-736, OM-735, OM-734, OM-723, OM-722, OM-719, OM-717, OM-715, OM-713, OM-710, OM-707, OM-705, OM-704, OM-702, OM-701, OM-699, OM-698, OM-697, OM-695, OM-694, OM-690, OM-688, OM-687, OM-686, OM-684, OM-683, OM-680, OM-679, OM-678, OM-677, OM-676, OM-670, OM-668, OM-665, OM-655, OM-654, OM-652, OM-651, OM-649, OM-647, OM-645, OM-644, OM-643, OM-642, OM-639, OM-638, OM-637, OM-636, OM-635, OM-634, OM-625) "
	jql = jql & " AND issuetype in (""Activity - Archive Validation"", ""Activity - Feasibility"", ""Activity - Manual Production"", ""Activity - Other"", ""Activity - Quality Analysis"", ""Activity - SQR Measurement"", ""Activity - Source Acquisition"", ""Activity - Source Acquisition - Field"", ""Activity - Source Analysis"", ""Activity - Source Preparation"")"
	jql = jql & " AND ""Assigned Unit"" in (""SO MOMA"",""SO SSO"",""SO APA"",""SO AME"",""SO PMO"",""SO GDT"",""SO EAP"", ""SO ECA"", ""SO SAMEA"",""SO EECA"", ""SO AFR"", ""SO WCE"", ""SO STS"", ""SO SAM"", ""SO PDV"", ""SO OCE"", ""SO NEA"", ""SO NAM"", ""SO LAM"")"
	jql = jql & " AND ("
	' ' jql = jql & " status changed to done during (""" & from_ & """, """ & to_ & """)"
	'jql = jql & " status changed to done during (""" & from_ & """, """ & to_ & """)"
	' jql = jql & " OR status changed to ""done"" during (""" & from_ & """, """ & to_ & """)"
	' 'jql = jql & " OR status changed to planned during (""" & from_ & """"", """ & to_ & """)"
	jql = jql & " ""Baseline end date"" >= """ & from_ & """ and ""Baseline end date"" <= """ & to_ & """"
	jql = jql & " )"
	jql = jql & " and ("
	jql = jql & " resolution in (""Accepted"",""Unresolved"",""Rejected"") "
	jql = jql & " ) "
	jql = jql & " and ("
	jql = jql & " status not in (""Open"",""On Hold"",""To Review"") "
	jql = jql & " ) "
'for testing
'jql = jql & " and key = OM-54565 "

	expl = ""
	expl = expl & "<span style='font-family:verdana;font-size:10pt;'><b>OTD Miss Rate</b></span><br>"
	expl = expl & "<span style='font-family:verdana;font-size:8pt;'>"
	expl = expl & "All issues of type (Activity - Archive Validation,  Activity - Feasibility,  Activity - Manual Production,  Activity - Other,  Activity - Quality Analysis,  Activity - SQR Measurement,  Activity - Source Acquisition,  Activity - Source Acquisition - Field,  Activity - Source Analysis,  Activity - Source Preparation)<br>"
	expl = expl & "and Assigned Unit of (SO MOMA, SO SSO, SO APA, SO AME, SO PMO, SO GDT, SO EAP, SO ECA, SO SAMEA, SO EECA,  SO AFR,  SO WCE,  SO STS,  SO SAM,  SO PDV,  SO OCE,  SO NEA,  SO NAM,  SO LAM)<br>"
	expl = expl & "and and Baseline End Date in month XX<br>"
	expl = expl & "and resolution in Accepted/Rejected/Unresolved.<br>"
	expl = expl & "and current state not in Open,On Hold and To Review<br>"
	expl = expl & "Review for missed End Date.<br>"
	expl = expl & "</span>"
end if


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
	response.write expl

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
	'response.write "<td>" & "state change overview"   & "</td>"
	response.write "<td>" & "date set to Done"   & "</td>"
	'response.write "<td>" & "date set to closed"   & "</td>"
	'response.write "<td>" & "created"   & "</td>"
	'response.write "<td>" & "updated"   & "</td>"
	' response.write "<td>" & "item duedate"   & "</td>"
	' response.write "<td>" & "item target duedate"   & "</td>"
	' response.write "<td>" & "duedate via linked sprint(s)"   & "</td>"
	response.write "<td>" & "OTD miss (baseline updated after planned state)" & "</td>"
	response.write "<td>" & "Change requested by" & "</td>"
	response.write "<td>" & "Change request type" & "</td>"
	response.write "<td>" & "Baseline enddate when state got planned" & "</td>"
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
			response.write "<td>" & getFieldValue(item_,"summary") & "</td>"
			response.write "<td>" & getFieldValue(item_,"status") & "</td>"
			response.write "<td>" & getFieldValue(item_,"assigned_unit") & "</td>"
			response.write "<td>" & getFieldValue(item_,"requesting_unit") & "</td>"
			response.write "<td>" & toDD_MM_YYYY(getFieldValue(item_,"bl_enddate")) & "</td>"
			response.write "<td>" & toDD_MM_YYYY(getFieldValue(item_,"enddate")) & "</td>"
			'response.write "<td>" & replace(getFieldValue(item_,"transitions_history"), "@@", "<br>") & "</td>"
			response.write "<td>" & toDD_MM_YYYY(getSettoDone(getFieldValue(item_,"transitions_history"))) & "</td>"
			'response.write "<td>" & toDD_MM_YYYY(getSettoClosed(getFieldValue(item_,"transitions_history"))) & "</td>"
			'response.write "<td>" & toDD_MM_YYYY(getFieldValue(item_,"created")) & "</td>"
			'response.write "<td>" & toDD_MM_YYYY(getFieldValue(item_,"updated")) & "</td>"

			'8JUN2020 - did we change baseline after item was set to planned
			response.write "<td>" & BL_changed_after_Planned(item_) & "</td>"
			response.write "<td>" & SO_comment_(item_) & "</td>"
			response.write "<td>" & SO_comment_type_(item_) & "</td>"
			'16JUL2020 - what was the baseline enddate when state=planned
			response.write "<td>" & BL_date_when_planned(item_) & "</td>"

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
function toDate_format(s)
toDate_format = s
if toDate_format <> "" then
	toDate_format = mid(s,7,2) & "/" & mid(s,5,2) & "/" & mid(s,1,4)
end if
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
'response.write "https://soreporting.azurewebsites.net/src_reporting/getJiraitems.aspx?project=" & project & "&yearmonth=" & yearmonth & "&rnd=" & rnd & ""
randomize timer
'On Error Resume Next
Dim xmlhttp
Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
' response.write  "https://soreporting.azurewebsites.net/so_reporting/getJiraitems.aspx?jql=" & jql & "&rnd=" & rnd & ""
' response.end
xmlhttp.Open "GET", "https://soreporting.azurewebsites.net/so_reporting/getJiraitems.aspx?jql=" & jql & "&rnd=" & rnd & "", False
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

function SO_comment_(x)
SO_comment_ = "none"
dim comm_hist
comm_hist = getFieldValue(x, "comments_history")
'do some simple changes so our pattern matches
comm_hist = lcase(comm_hist) 'bring to lowercase
comm_hist = replace(comm_hist, " ", "") 'remove all spaces
'becuase the text contain RTF conctrol chars we cannot look for the whoe string, instead we'll look for the different words and make sure they appear after eachother
dim arr
dim a
dim arr2

dim p1
dim p2
dim p3

arr = split(comm_hist, "@@")
for a = lbound(arr) to ubound(arr)
if arr(a) <> "" then

	arr2 = split(arr(a), "||")
	p1 = 0
	p2 = 0
	p3 = 0

	p1 = InStr(arr2(1), "type")
	If p1 <> 0 Then p2 = InStr(p1, arr2(1), "initiator")
	If p2 <> 0 Then p3 = InStr(p2, arr2(1), "so")

	If p1 < p2 And p2 < p3 Then 'ok they appear in this order
	If p3 - p1 <= 100 Then 'ok and they are pretty close to each other (char range 100 chars)
		SO_comment_ = "SO"
	End If
	End If

	p1 = InStr(arr2(1), "type")
	If p1 <> 0 Then p2 = InStr(p1, arr2(1), "initiator")
	If p2 <> 0 Then p3 = InStr(p2, arr2(1), "pu")

	If p1 < p2 And p2 < p3 Then 'ok they appear in this order
	If p3 - p1 <= 100 Then 'ok and they are pretty close to each other (char range 100 chars)
		SO_comment_ = "PU"
	End If
	End If

end if
next
end function

function BL_date_when_planned(x)
dim bl_hist
dim trn_hist

bl_hist = getFieldValue(x, "bl_enddate_history")
trn_hist = getFieldValue(x, "transitions_history")

dim arr
dim arr2
dim a
dim a2

dim planned_
planned_ = "00000000"

'first check when planned was set the first time
arr = split(trn_hist, "@@")
for a = lbound(arr) to ubound(arr)
if arr(a) <> "" then
	arr2 = split(arr(a), "||")
	if arr2(2) = "Planned" then
		planned_ = arr2(0)
		a = ubound(arr)
	end if
end if
next

if planned_ <> "00000000" then
If bl_hist <> "" Then

	BL_date_when_planned = ""
    arr = Split(bl_hist, "@@")
    For a = UBound(arr) To LBound(arr) Step -1
    If arr(a) <> "" Then
        arr2 = Split(arr(a), "||")
        If planned_ >= arr2(0) Then
            BL_date_when_planned = Replace(arr2(2), "-", "")
            a = -1
        Else
        End If
    End If
    Next

    If BL_date_when_planned = "" Then
        'the date must have been the date when set at createion
        arr = Split(bl_hist, "@@")
        If arr(0) <> "" Then
            arr2 = Split(arr(0), "||")
            BL_date_when_planned = Replace(arr2(1), "-", "")
        End If
    End If

end if
end if
end function

function BL_changed_after_Planned(x)
BL_changed_after_Planned = false
'exit function

dim bl_hist
dim trn_hist

bl_hist = getFieldValue(x, "bl_enddate_history")
trn_hist = getFieldValue(x, "transitions_history")

dim arr
dim arr2
dim a
dim a2

dim planned_
planned_ = "00000000"

'first check when planned was set the first time
arr = split(trn_hist, "@@")
for a = lbound(arr) to ubound(arr)
if arr(a) <> "" then
	arr2 = split(arr(a), "||")
	if arr2(2) = "Planned" then
		planned_ = arr2(0)
		a = ubound(arr)
	end if
end if
next

if planned_ <> "00000000" then
	'then check if we have a bl_enddate change after this date
	arr = split(bl_hist, "@@")
	for a = lbound(arr) to ubound(arr)
	if arr(a) <> "" then
		arr2 = split(arr(a), "||")
		if arr2(0) > planned_ then
			BL_changed_after_Planned = true
			a = ubound(arr)
		end if
	end if
	next
end if

end function


function SO_comment_type_(x)
SO_comment_type_ = ""
dim comm_hist
comm_hist = getFieldValue(x, "comments_history")
'do some simple changes so our pattern matches
comm_hist = lcase(comm_hist) 'bring to lowercase
comm_hist = replace(comm_hist, " ", "") 'remove all spaces
'becuase the text contain RTF conctrol chars we cannot look for the whoe string, instead we'll look for the different words and make sure they appear after eachother
dim arr
dim a
dim arr2

dim p1
dim p2
dim p3

arr = split(comm_hist, "@@")
for a = lbound(arr) to ubound(arr)
if arr(a) <> "" then

	arr2 = split(arr(a), "||")
	p1 = 0
	p2 = 0
	p3 = 0

	p1 = InStr(arr2(1), "type")
	If p1 <> 0 Then p2 = InStr(p1, arr2(1), "time")
	If p2 <> 0 Then p3 = InStr(p2, arr2(1), "initiator")

	If p1 < p2 And p2 < p3 Then 'ok they appear in this order
	If p3 - p1 <= 100 Then 'ok and they are pretty close to each other (char range 100 chars)
		SO_comment_type_ = "time"
	End If
	End If

	p1 = InStr(arr2(1), "type")
	If p1 <> 0 Then p2 = InStr(p1, arr2(1), "scope")
	If p2 <> 0 Then p3 = InStr(p2, arr2(1), "initiator")

	If p1 < p2 And p2 < p3 Then 'ok they appear in this order
	If p3 - p1 <= 100 Then 'ok and they are pretty close to each other (char range 100 chars)
		SO_comment_type_ = "scope"
	End If
	End If

end if
next
end function

%>