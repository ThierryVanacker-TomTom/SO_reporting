<%@ Page Language="VB" Explicit="True" Debug="true"%>

<%@ Import Namespace="System" %>
<%@ Import Namespace="System.Collections" %>
<%@ Import Namespace="System.Web.Services" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OleDb" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Data.Linq" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="System.Text" %>
<%@ Import Namespace="System.Text.RegularExpressions" %>
<%@ Import Namespace="DocumentFormat.OpenXml" %>
<%@ Import Namespace="DocumentFormat.OpenXml.Packaging" %>
<%@ Import Namespace="DocumentFormat.OpenXml.Spreadsheet" %>
<%@ Import Namespace="DocumentFormat.OpenXml.Wordprocessing" %>
<%@ Import Namespace="System.Web.Script.Serialization" %>

<script language=vb runat=server>
Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs)
System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12

response.write ("RM-302;Start date:" & parseHist_("RM-302") & ";" & "End date:" & parseHist2_("RM-302") & "<br>" )
response.write ("RM-298;Start date:" & parseHist_("RM-298") & ";" & "End date:" & parseHist2_("RM-298") & "<br>" )
response.write ("RM-296;Start date:" & parseHist_("RM-296") & ";" & "End date:" & parseHist2_("RM-296") & "<br>" )
response.write ("RM-295;Start date:" & parseHist_("RM-295") & ";" & "End date:" & parseHist2_("RM-295") & "<br>" )
response.write ("RM-294;Start date:" & parseHist_("RM-294") & ";" & "End date:" & parseHist2_("RM-294") & "<br>" )
response.write ("RM-287;Start date:" & parseHist_("RM-287") & ";" & "End date:" & parseHist2_("RM-287") & "<br>" )
response.write ("RM-282;Start date:" & parseHist_("RM-282") & ";" & "End date:" & parseHist2_("RM-282") & "<br>" )
response.write ("RM-281;Start date:" & parseHist_("RM-281") & ";" & "End date:" & parseHist2_("RM-281") & "<br>" )
response.write ("RM-270;Start date:" & parseHist_("RM-270") & ";" & "End date:" & parseHist2_("RM-270") & "<br>" )
response.write ("RM-268;Start date:" & parseHist_("RM-268") & ";" & "End date:" & parseHist2_("RM-268") & "<br>" )
response.write ("RM-244;Start date:" & parseHist_("RM-244") & ";" & "End date:" & parseHist2_("RM-244") & "<br>" )
response.write ("RM-243;Start date:" & parseHist_("RM-243") & ";" & "End date:" & parseHist2_("RM-243") & "<br>" )
response.write ("RM-241;Start date:" & parseHist_("RM-241") & ";" & "End date:" & parseHist2_("RM-241") & "<br>" )
response.write ("RM-240;Start date:" & parseHist_("RM-240") & ";" & "End date:" & parseHist2_("RM-240") & "<br>" )
response.write ("RM-239;Start date:" & parseHist_("RM-239") & ";" & "End date:" & parseHist2_("RM-239") & "<br>" )
response.write ("RM-238;Start date:" & parseHist_("RM-238") & ";" & "End date:" & parseHist2_("RM-238") & "<br>" )
response.write ("RM-237;Start date:" & parseHist_("RM-237") & ";" & "End date:" & parseHist2_("RM-237") & "<br>" )
response.write ("RM-236;Start date:" & parseHist_("RM-236") & ";" & "End date:" & parseHist2_("RM-236") & "<br>" )
response.write ("RM-235;Start date:" & parseHist_("RM-235") & ";" & "End date:" & parseHist2_("RM-235") & "<br>" )
response.write ("RM-234;Start date:" & parseHist_("RM-234") & ";" & "End date:" & parseHist2_("RM-234") & "<br>" )
response.write ("RM-233;Start date:" & parseHist_("RM-233") & ";" & "End date:" & parseHist2_("RM-233") & "<br>" )
response.write ("RM-232;Start date:" & parseHist_("RM-232") & ";" & "End date:" & parseHist2_("RM-232") & "<br>" )
response.write ("RM-231;Start date:" & parseHist_("RM-231") & ";" & "End date:" & parseHist2_("RM-231") & "<br>" )
response.write ("RM-230;Start date:" & parseHist_("RM-230") & ";" & "End date:" & parseHist2_("RM-230") & "<br>" )
response.write ("RM-227;Start date:" & parseHist_("RM-227") & ";" & "End date:" & parseHist2_("RM-227") & "<br>" )
response.write ("RM-225;Start date:" & parseHist_("RM-225") & ";" & "End date:" & parseHist2_("RM-225") & "<br>" )
response.write ("RM-224;Start date:" & parseHist_("RM-224") & ";" & "End date:" & parseHist2_("RM-224") & "<br>" )
response.write ("RM-223;Start date:" & parseHist_("RM-223") & ";" & "End date:" & parseHist2_("RM-223") & "<br>" )
response.write ("RM-222;Start date:" & parseHist_("RM-222") & ";" & "End date:" & parseHist2_("RM-222") & "<br>" )
response.write ("RM-218;Start date:" & parseHist_("RM-218") & ";" & "End date:" & parseHist2_("RM-218") & "<br>" )
response.write ("RM-200;Start date:" & parseHist_("RM-200") & ";" & "End date:" & parseHist2_("RM-200") & "<br>" )
response.write ("OM-19383;Start date:" & parseHist_("OM-19383") & ";" & "End date:" & parseHist2_("OM-19383") & "<br>" )
response.write ("OM-19314;Start date:" & parseHist_("OM-19314") & ";" & "End date:" & parseHist2_("OM-19314") & "<br>" )
response.write ("OM-19194;Start date:" & parseHist_("OM-19194") & ";" & "End date:" & parseHist2_("OM-19194") & "<br>" )
response.write ("OM-18907;Start date:" & parseHist_("OM-18907") & ";" & "End date:" & parseHist2_("OM-18907") & "<br>" )
response.write ("OM-18552;Start date:" & parseHist_("OM-18552") & ";" & "End date:" & parseHist2_("OM-18552") & "<br>" )
response.write ("OM-18480;Start date:" & parseHist_("OM-18480") & ";" & "End date:" & parseHist2_("OM-18480") & "<br>" )
response.write ("OM-18475;Start date:" & parseHist_("OM-18475") & ";" & "End date:" & parseHist2_("OM-18475") & "<br>" )
response.write ("OM-18208;Start date:" & parseHist_("OM-18208") & ";" & "End date:" & parseHist2_("OM-18208") & "<br>" )
response.write ("OM-17985;Start date:" & parseHist_("OM-17985") & ";" & "End date:" & parseHist2_("OM-17985") & "<br>" )
response.write ("OM-17948;Start date:" & parseHist_("OM-17948") & ";" & "End date:" & parseHist2_("OM-17948") & "<br>" )
response.write ("OM-17821;Start date:" & parseHist_("OM-17821") & ";" & "End date:" & parseHist2_("OM-17821") & "<br>" )
response.write ("OM-17809;Start date:" & parseHist_("OM-17809") & ";" & "End date:" & parseHist2_("OM-17809") & "<br>" )
response.write ("OM-17808;Start date:" & parseHist_("OM-17808") & ";" & "End date:" & parseHist2_("OM-17808") & "<br>" )
response.write ("OM-17788;Start date:" & parseHist_("OM-17788") & ";" & "End date:" & parseHist2_("OM-17788") & "<br>" )
response.write ("OM-17731;Start date:" & parseHist_("OM-17731") & ";" & "End date:" & parseHist2_("OM-17731") & "<br>" )
response.write ("OM-17646;Start date:" & parseHist_("OM-17646") & ";" & "End date:" & parseHist2_("OM-17646") & "<br>" )
response.write ("OM-17580;Start date:" & parseHist_("OM-17580") & ";" & "End date:" & parseHist2_("OM-17580") & "<br>" )
response.write ("OM-17557;Start date:" & parseHist_("OM-17557") & ";" & "End date:" & parseHist2_("OM-17557") & "<br>" )
response.write ("OM-17532;Start date:" & parseHist_("OM-17532") & ";" & "End date:" & parseHist2_("OM-17532") & "<br>" )
response.write ("OM-17362;Start date:" & parseHist_("OM-17362") & ";" & "End date:" & parseHist2_("OM-17362") & "<br>" )
response.write ("OM-17360;Start date:" & parseHist_("OM-17360") & ";" & "End date:" & parseHist2_("OM-17360") & "<br>" )
response.write ("OM-17331;Start date:" & parseHist_("OM-17331") & ";" & "End date:" & parseHist2_("OM-17331") & "<br>" )
response.write ("OM-16956;Start date:" & parseHist_("OM-16956") & ";" & "End date:" & parseHist2_("OM-16956") & "<br>" )
response.write ("OM-16936;Start date:" & parseHist_("OM-16936") & ";" & "End date:" & parseHist2_("OM-16936") & "<br>" )
response.write ("OM-16846;Start date:" & parseHist_("OM-16846") & ";" & "End date:" & parseHist2_("OM-16846") & "<br>" )
response.write ("OM-16826;Start date:" & parseHist_("OM-16826") & ";" & "End date:" & parseHist2_("OM-16826") & "<br>" )
response.write ("OM-16813;Start date:" & parseHist_("OM-16813") & ";" & "End date:" & parseHist2_("OM-16813") & "<br>" )
response.write ("OM-16698;Start date:" & parseHist_("OM-16698") & ";" & "End date:" & parseHist2_("OM-16698") & "<br>" )
response.write ("OM-16690;Start date:" & parseHist_("OM-16690") & ";" & "End date:" & parseHist2_("OM-16690") & "<br>" )
response.write ("OM-16689;Start date:" & parseHist_("OM-16689") & ";" & "End date:" & parseHist2_("OM-16689") & "<br>" )
response.write ("OM-16688;Start date:" & parseHist_("OM-16688") & ";" & "End date:" & parseHist2_("OM-16688") & "<br>" )
response.write ("OM-16676;Start date:" & parseHist_("OM-16676") & ";" & "End date:" & parseHist2_("OM-16676") & "<br>" )
response.write ("OM-16633;Start date:" & parseHist_("OM-16633") & ";" & "End date:" & parseHist2_("OM-16633") & "<br>" )
response.write ("OM-16600;Start date:" & parseHist_("OM-16600") & ";" & "End date:" & parseHist2_("OM-16600") & "<br>" )
response.write ("OM-16575;Start date:" & parseHist_("OM-16575") & ";" & "End date:" & parseHist2_("OM-16575") & "<br>" )
response.write ("OM-16489;Start date:" & parseHist_("OM-16489") & ";" & "End date:" & parseHist2_("OM-16489") & "<br>" )
response.write ("OM-16478;Start date:" & parseHist_("OM-16478") & ";" & "End date:" & parseHist2_("OM-16478") & "<br>" )
response.write ("OM-16475;Start date:" & parseHist_("OM-16475") & ";" & "End date:" & parseHist2_("OM-16475") & "<br>" )
response.write ("OM-16474;Start date:" & parseHist_("OM-16474") & ";" & "End date:" & parseHist2_("OM-16474") & "<br>" )
response.write ("OM-16473;Start date:" & parseHist_("OM-16473") & ";" & "End date:" & parseHist2_("OM-16473") & "<br>" )
response.write ("OM-16468;Start date:" & parseHist_("OM-16468") & ";" & "End date:" & parseHist2_("OM-16468") & "<br>" )
response.write ("OM-16467;Start date:" & parseHist_("OM-16467") & ";" & "End date:" & parseHist2_("OM-16467") & "<br>" )
response.write ("OM-16465;Start date:" & parseHist_("OM-16465") & ";" & "End date:" & parseHist2_("OM-16465") & "<br>" )
response.write ("OM-16390;Start date:" & parseHist_("OM-16390") & ";" & "End date:" & parseHist2_("OM-16390") & "<br>" )
response.write ("OM-16363;Start date:" & parseHist_("OM-16363") & ";" & "End date:" & parseHist2_("OM-16363") & "<br>" )
response.write ("OM-16362;Start date:" & parseHist_("OM-16362") & ";" & "End date:" & parseHist2_("OM-16362") & "<br>" )
response.write ("OM-16330;Start date:" & parseHist_("OM-16330") & ";" & "End date:" & parseHist2_("OM-16330") & "<br>" )
response.write ("OM-16297;Start date:" & parseHist_("OM-16297") & ";" & "End date:" & parseHist2_("OM-16297") & "<br>" )
response.write ("OM-16276;Start date:" & parseHist_("OM-16276") & ";" & "End date:" & parseHist2_("OM-16276") & "<br>" )
response.write ("OM-16247;Start date:" & parseHist_("OM-16247") & ";" & "End date:" & parseHist2_("OM-16247") & "<br>" )
response.write ("OM-16192;Start date:" & parseHist_("OM-16192") & ";" & "End date:" & parseHist2_("OM-16192") & "<br>" )
response.write ("OM-16185;Start date:" & parseHist_("OM-16185") & ";" & "End date:" & parseHist2_("OM-16185") & "<br>" )
response.write ("OM-16180;Start date:" & parseHist_("OM-16180") & ";" & "End date:" & parseHist2_("OM-16180") & "<br>" )
response.write ("OM-16109;Start date:" & parseHist_("OM-16109") & ";" & "End date:" & parseHist2_("OM-16109") & "<br>" )
response.write ("OM-16053;Start date:" & parseHist_("OM-16053") & ";" & "End date:" & parseHist2_("OM-16053") & "<br>" )
response.write ("OM-16051;Start date:" & parseHist_("OM-16051") & ";" & "End date:" & parseHist2_("OM-16051") & "<br>" )
response.write ("OM-16015;Start date:" & parseHist_("OM-16015") & ";" & "End date:" & parseHist2_("OM-16015") & "<br>" )
response.write ("OM-16005;Start date:" & parseHist_("OM-16005") & ";" & "End date:" & parseHist2_("OM-16005") & "<br>" )
response.write ("OM-15974;Start date:" & parseHist_("OM-15974") & ";" & "End date:" & parseHist2_("OM-15974") & "<br>" )
response.write ("OM-15795;Start date:" & parseHist_("OM-15795") & ";" & "End date:" & parseHist2_("OM-15795") & "<br>" )
response.write ("OM-15784;Start date:" & parseHist_("OM-15784") & ";" & "End date:" & parseHist2_("OM-15784") & "<br>" )
response.write ("OM-15783;Start date:" & parseHist_("OM-15783") & ";" & "End date:" & parseHist2_("OM-15783") & "<br>" )
response.write ("OM-15782;Start date:" & parseHist_("OM-15782") & ";" & "End date:" & parseHist2_("OM-15782") & "<br>" )
response.write ("OM-15781;Start date:" & parseHist_("OM-15781") & ";" & "End date:" & parseHist2_("OM-15781") & "<br>" )
response.write ("OM-15733;Start date:" & parseHist_("OM-15733") & ";" & "End date:" & parseHist2_("OM-15733") & "<br>" )
response.write ("OM-15722;Start date:" & parseHist_("OM-15722") & ";" & "End date:" & parseHist2_("OM-15722") & "<br>" )
response.write ("OM-15638;Start date:" & parseHist_("OM-15638") & ";" & "End date:" & parseHist2_("OM-15638") & "<br>" )
response.write ("OM-15594;Start date:" & parseHist_("OM-15594") & ";" & "End date:" & parseHist2_("OM-15594") & "<br>" )
response.write ("OM-15576;Start date:" & parseHist_("OM-15576") & ";" & "End date:" & parseHist2_("OM-15576") & "<br>" )
response.write ("OM-15380;Start date:" & parseHist_("OM-15380") & ";" & "End date:" & parseHist2_("OM-15380") & "<br>" )
response.write ("OM-15327;Start date:" & parseHist_("OM-15327") & ";" & "End date:" & parseHist2_("OM-15327") & "<br>" )
response.write ("OM-15326;Start date:" & parseHist_("OM-15326") & ";" & "End date:" & parseHist2_("OM-15326") & "<br>" )
response.write ("OM-15323;Start date:" & parseHist_("OM-15323") & ";" & "End date:" & parseHist2_("OM-15323") & "<br>" )
response.write ("OM-15303;Start date:" & parseHist_("OM-15303") & ";" & "End date:" & parseHist2_("OM-15303") & "<br>" )
response.write ("OM-15207;Start date:" & parseHist_("OM-15207") & ";" & "End date:" & parseHist2_("OM-15207") & "<br>" )
response.write ("OM-15203;Start date:" & parseHist_("OM-15203") & ";" & "End date:" & parseHist2_("OM-15203") & "<br>" )
response.write ("OM-15199;Start date:" & parseHist_("OM-15199") & ";" & "End date:" & parseHist2_("OM-15199") & "<br>" )
response.write ("OM-15190;Start date:" & parseHist_("OM-15190") & ";" & "End date:" & parseHist2_("OM-15190") & "<br>" )
response.write ("OM-15170;Start date:" & parseHist_("OM-15170") & ";" & "End date:" & parseHist2_("OM-15170") & "<br>" )
response.write ("OM-15169;Start date:" & parseHist_("OM-15169") & ";" & "End date:" & parseHist2_("OM-15169") & "<br>" )
response.write ("OM-15166;Start date:" & parseHist_("OM-15166") & ";" & "End date:" & parseHist2_("OM-15166") & "<br>" )
response.write ("OM-15147;Start date:" & parseHist_("OM-15147") & ";" & "End date:" & parseHist2_("OM-15147") & "<br>" )
response.write ("OM-15133;Start date:" & parseHist_("OM-15133") & ";" & "End date:" & parseHist2_("OM-15133") & "<br>" )
response.write ("OM-15128;Start date:" & parseHist_("OM-15128") & ";" & "End date:" & parseHist2_("OM-15128") & "<br>" )
response.write ("OM-15080;Start date:" & parseHist_("OM-15080") & ";" & "End date:" & parseHist2_("OM-15080") & "<br>" )
response.write ("OM-15075;Start date:" & parseHist_("OM-15075") & ";" & "End date:" & parseHist2_("OM-15075") & "<br>" )
response.write ("OM-15055;Start date:" & parseHist_("OM-15055") & ";" & "End date:" & parseHist2_("OM-15055") & "<br>" )
response.write ("OM-15039;Start date:" & parseHist_("OM-15039") & ";" & "End date:" & parseHist2_("OM-15039") & "<br>" )
response.write ("OM-15038;Start date:" & parseHist_("OM-15038") & ";" & "End date:" & parseHist2_("OM-15038") & "<br>" )
response.write ("OM-15035;Start date:" & parseHist_("OM-15035") & ";" & "End date:" & parseHist2_("OM-15035") & "<br>" )
response.write ("OM-15032;Start date:" & parseHist_("OM-15032") & ";" & "End date:" & parseHist2_("OM-15032") & "<br>" )
response.write ("OM-15003;Start date:" & parseHist_("OM-15003") & ";" & "End date:" & parseHist2_("OM-15003") & "<br>" )
response.write ("OM-14999;Start date:" & parseHist_("OM-14999") & ";" & "End date:" & parseHist2_("OM-14999") & "<br>" )
response.write ("OM-14998;Start date:" & parseHist_("OM-14998") & ";" & "End date:" & parseHist2_("OM-14998") & "<br>" )
response.write ("OM-14997;Start date:" & parseHist_("OM-14997") & ";" & "End date:" & parseHist2_("OM-14997") & "<br>" )
response.write ("OM-14964;Start date:" & parseHist_("OM-14964") & ";" & "End date:" & parseHist2_("OM-14964") & "<br>" )
response.write ("OM-14958;Start date:" & parseHist_("OM-14958") & ";" & "End date:" & parseHist2_("OM-14958") & "<br>" )
response.write ("OM-14957;Start date:" & parseHist_("OM-14957") & ";" & "End date:" & parseHist2_("OM-14957") & "<br>" )
response.write ("OM-14943;Start date:" & parseHist_("OM-14943") & ";" & "End date:" & parseHist2_("OM-14943") & "<br>" )
response.write ("OM-14864;Start date:" & parseHist_("OM-14864") & ";" & "End date:" & parseHist2_("OM-14864") & "<br>" )
response.write ("OM-14798;Start date:" & parseHist_("OM-14798") & ";" & "End date:" & parseHist2_("OM-14798") & "<br>" )
response.write ("OM-14790;Start date:" & parseHist_("OM-14790") & ";" & "End date:" & parseHist2_("OM-14790") & "<br>" )
response.write ("OM-14784;Start date:" & parseHist_("OM-14784") & ";" & "End date:" & parseHist2_("OM-14784") & "<br>" )
response.write ("OM-14768;Start date:" & parseHist_("OM-14768") & ";" & "End date:" & parseHist2_("OM-14768") & "<br>" )
response.write ("OM-14767;Start date:" & parseHist_("OM-14767") & ";" & "End date:" & parseHist2_("OM-14767") & "<br>" )
response.write ("OM-14765;Start date:" & parseHist_("OM-14765") & ";" & "End date:" & parseHist2_("OM-14765") & "<br>" )
response.write ("OM-14736;Start date:" & parseHist_("OM-14736") & ";" & "End date:" & parseHist2_("OM-14736") & "<br>" )
response.write ("OM-14706;Start date:" & parseHist_("OM-14706") & ";" & "End date:" & parseHist2_("OM-14706") & "<br>" )
response.write ("OM-14695;Start date:" & parseHist_("OM-14695") & ";" & "End date:" & parseHist2_("OM-14695") & "<br>" )
response.write ("OM-14671;Start date:" & parseHist_("OM-14671") & ";" & "End date:" & parseHist2_("OM-14671") & "<br>" )
response.write ("OM-14664;Start date:" & parseHist_("OM-14664") & ";" & "End date:" & parseHist2_("OM-14664") & "<br>" )
response.write ("OM-14663;Start date:" & parseHist_("OM-14663") & ";" & "End date:" & parseHist2_("OM-14663") & "<br>" )
response.write ("OM-14660;Start date:" & parseHist_("OM-14660") & ";" & "End date:" & parseHist2_("OM-14660") & "<br>" )
response.write ("OM-14542;Start date:" & parseHist_("OM-14542") & ";" & "End date:" & parseHist2_("OM-14542") & "<br>" )
response.write ("OM-14431;Start date:" & parseHist_("OM-14431") & ";" & "End date:" & parseHist2_("OM-14431") & "<br>" )
response.write ("OM-14408;Start date:" & parseHist_("OM-14408") & ";" & "End date:" & parseHist2_("OM-14408") & "<br>" )
response.write ("OM-14214;Start date:" & parseHist_("OM-14214") & ";" & "End date:" & parseHist2_("OM-14214") & "<br>" )
response.write ("OM-14073;Start date:" & parseHist_("OM-14073") & ";" & "End date:" & parseHist2_("OM-14073") & "<br>" )
response.write ("OM-14066;Start date:" & parseHist_("OM-14066") & ";" & "End date:" & parseHist2_("OM-14066") & "<br>" )
response.write ("OM-14050;Start date:" & parseHist_("OM-14050") & ";" & "End date:" & parseHist2_("OM-14050") & "<br>" )
response.write ("OM-13979;Start date:" & parseHist_("OM-13979") & ";" & "End date:" & parseHist2_("OM-13979") & "<br>" )
response.write ("OM-13977;Start date:" & parseHist_("OM-13977") & ";" & "End date:" & parseHist2_("OM-13977") & "<br>" )
response.write ("OM-13856;Start date:" & parseHist_("OM-13856") & ";" & "End date:" & parseHist2_("OM-13856") & "<br>" )
response.write ("OM-13853;Start date:" & parseHist_("OM-13853") & ";" & "End date:" & parseHist2_("OM-13853") & "<br>" )
response.write ("OM-13852;Start date:" & parseHist_("OM-13852") & ";" & "End date:" & parseHist2_("OM-13852") & "<br>" )
response.write ("OM-13850;Start date:" & parseHist_("OM-13850") & ";" & "End date:" & parseHist2_("OM-13850") & "<br>" )
response.write ("OM-13849;Start date:" & parseHist_("OM-13849") & ";" & "End date:" & parseHist2_("OM-13849") & "<br>" )
response.write ("OM-13846;Start date:" & parseHist_("OM-13846") & ";" & "End date:" & parseHist2_("OM-13846") & "<br>" )
response.write ("OM-13810;Start date:" & parseHist_("OM-13810") & ";" & "End date:" & parseHist2_("OM-13810") & "<br>" )
response.write ("OM-13808;Start date:" & parseHist_("OM-13808") & ";" & "End date:" & parseHist2_("OM-13808") & "<br>" )
response.write ("OM-13581;Start date:" & parseHist_("OM-13581") & ";" & "End date:" & parseHist2_("OM-13581") & "<br>" )
response.write ("OM-13543;Start date:" & parseHist_("OM-13543") & ";" & "End date:" & parseHist2_("OM-13543") & "<br>" )
response.write ("OM-13538;Start date:" & parseHist_("OM-13538") & ";" & "End date:" & parseHist2_("OM-13538") & "<br>" )
response.write ("OM-13509;Start date:" & parseHist_("OM-13509") & ";" & "End date:" & parseHist2_("OM-13509") & "<br>" )
response.write ("OM-13470;Start date:" & parseHist_("OM-13470") & ";" & "End date:" & parseHist2_("OM-13470") & "<br>" )
response.write ("OM-13375;Start date:" & parseHist_("OM-13375") & ";" & "End date:" & parseHist2_("OM-13375") & "<br>" )
response.write ("OM-13293;Start date:" & parseHist_("OM-13293") & ";" & "End date:" & parseHist2_("OM-13293") & "<br>" )
response.write ("OM-13292;Start date:" & parseHist_("OM-13292") & ";" & "End date:" & parseHist2_("OM-13292") & "<br>" )
response.write ("OM-13291;Start date:" & parseHist_("OM-13291") & ";" & "End date:" & parseHist2_("OM-13291") & "<br>" )
response.write ("OM-13271;Start date:" & parseHist_("OM-13271") & ";" & "End date:" & parseHist2_("OM-13271") & "<br>" )
response.write ("OM-13267;Start date:" & parseHist_("OM-13267") & ";" & "End date:" & parseHist2_("OM-13267") & "<br>" )
response.write ("OM-13261;Start date:" & parseHist_("OM-13261") & ";" & "End date:" & parseHist2_("OM-13261") & "<br>" )
response.write ("OM-13180;Start date:" & parseHist_("OM-13180") & ";" & "End date:" & parseHist2_("OM-13180") & "<br>" )
response.write ("OM-13156;Start date:" & parseHist_("OM-13156") & ";" & "End date:" & parseHist2_("OM-13156") & "<br>" )
response.write ("OM-13149;Start date:" & parseHist_("OM-13149") & ";" & "End date:" & parseHist2_("OM-13149") & "<br>" )
response.write ("OM-13076;Start date:" & parseHist_("OM-13076") & ";" & "End date:" & parseHist2_("OM-13076") & "<br>" )
response.write ("OM-13057;Start date:" & parseHist_("OM-13057") & ";" & "End date:" & parseHist2_("OM-13057") & "<br>" )
response.write ("OM-13049;Start date:" & parseHist_("OM-13049") & ";" & "End date:" & parseHist2_("OM-13049") & "<br>" )
response.write ("OM-13008;Start date:" & parseHist_("OM-13008") & ";" & "End date:" & parseHist2_("OM-13008") & "<br>" )
response.write ("OM-12997;Start date:" & parseHist_("OM-12997") & ";" & "End date:" & parseHist2_("OM-12997") & "<br>" )
response.write ("OM-12989;Start date:" & parseHist_("OM-12989") & ";" & "End date:" & parseHist2_("OM-12989") & "<br>" )
response.write ("OM-12985;Start date:" & parseHist_("OM-12985") & ";" & "End date:" & parseHist2_("OM-12985") & "<br>" )
response.write ("OM-12938;Start date:" & parseHist_("OM-12938") & ";" & "End date:" & parseHist2_("OM-12938") & "<br>" )
response.write ("OM-12929;Start date:" & parseHist_("OM-12929") & ";" & "End date:" & parseHist2_("OM-12929") & "<br>" )
response.write ("OM-12928;Start date:" & parseHist_("OM-12928") & ";" & "End date:" & parseHist2_("OM-12928") & "<br>" )
response.write ("OM-12829;Start date:" & parseHist_("OM-12829") & ";" & "End date:" & parseHist2_("OM-12829") & "<br>" )
response.write ("OM-12763;Start date:" & parseHist_("OM-12763") & ";" & "End date:" & parseHist2_("OM-12763") & "<br>" )
response.write ("OM-12751;Start date:" & parseHist_("OM-12751") & ";" & "End date:" & parseHist2_("OM-12751") & "<br>" )
response.write ("OM-12717;Start date:" & parseHist_("OM-12717") & ";" & "End date:" & parseHist2_("OM-12717") & "<br>" )
response.write ("OM-12709;Start date:" & parseHist_("OM-12709") & ";" & "End date:" & parseHist2_("OM-12709") & "<br>" )
response.write ("OM-12708;Start date:" & parseHist_("OM-12708") & ";" & "End date:" & parseHist2_("OM-12708") & "<br>" )
response.write ("OM-12707;Start date:" & parseHist_("OM-12707") & ";" & "End date:" & parseHist2_("OM-12707") & "<br>" )
response.write ("OM-12571;Start date:" & parseHist_("OM-12571") & ";" & "End date:" & parseHist2_("OM-12571") & "<br>" )
response.write ("OM-12546;Start date:" & parseHist_("OM-12546") & ";" & "End date:" & parseHist2_("OM-12546") & "<br>" )
response.write ("OM-12545;Start date:" & parseHist_("OM-12545") & ";" & "End date:" & parseHist2_("OM-12545") & "<br>" )
response.write ("OM-12485;Start date:" & parseHist_("OM-12485") & ";" & "End date:" & parseHist2_("OM-12485") & "<br>" )
response.write ("OM-12483;Start date:" & parseHist_("OM-12483") & ";" & "End date:" & parseHist2_("OM-12483") & "<br>" )
response.write ("OM-12372;Start date:" & parseHist_("OM-12372") & ";" & "End date:" & parseHist2_("OM-12372") & "<br>" )
response.write ("OM-12363;Start date:" & parseHist_("OM-12363") & ";" & "End date:" & parseHist2_("OM-12363") & "<br>" )
response.write ("OM-12305;Start date:" & parseHist_("OM-12305") & ";" & "End date:" & parseHist2_("OM-12305") & "<br>" )
response.write ("OM-12290;Start date:" & parseHist_("OM-12290") & ";" & "End date:" & parseHist2_("OM-12290") & "<br>" )
response.write ("OM-12250;Start date:" & parseHist_("OM-12250") & ";" & "End date:" & parseHist2_("OM-12250") & "<br>" )
response.write ("OM-12218;Start date:" & parseHist_("OM-12218") & ";" & "End date:" & parseHist2_("OM-12218") & "<br>" )
response.write ("OM-12212;Start date:" & parseHist_("OM-12212") & ";" & "End date:" & parseHist2_("OM-12212") & "<br>" )
response.write ("OM-12138;Start date:" & parseHist_("OM-12138") & ";" & "End date:" & parseHist2_("OM-12138") & "<br>" )
response.write ("OM-11968;Start date:" & parseHist_("OM-11968") & ";" & "End date:" & parseHist2_("OM-11968") & "<br>" )
response.write ("OM-11923;Start date:" & parseHist_("OM-11923") & ";" & "End date:" & parseHist2_("OM-11923") & "<br>" )
response.write ("OM-11921;Start date:" & parseHist_("OM-11921") & ";" & "End date:" & parseHist2_("OM-11921") & "<br>" )
response.write ("OM-11919;Start date:" & parseHist_("OM-11919") & ";" & "End date:" & parseHist2_("OM-11919") & "<br>" )
response.write ("OM-11902;Start date:" & parseHist_("OM-11902") & ";" & "End date:" & parseHist2_("OM-11902") & "<br>" )
response.write ("OM-11898;Start date:" & parseHist_("OM-11898") & ";" & "End date:" & parseHist2_("OM-11898") & "<br>" )
response.write ("OM-11890;Start date:" & parseHist_("OM-11890") & ";" & "End date:" & parseHist2_("OM-11890") & "<br>" )
response.write ("OM-11831;Start date:" & parseHist_("OM-11831") & ";" & "End date:" & parseHist2_("OM-11831") & "<br>" )
response.write ("OM-11829;Start date:" & parseHist_("OM-11829") & ";" & "End date:" & parseHist2_("OM-11829") & "<br>" )
response.write ("OM-11825;Start date:" & parseHist_("OM-11825") & ";" & "End date:" & parseHist2_("OM-11825") & "<br>" )
response.write ("OM-11821;Start date:" & parseHist_("OM-11821") & ";" & "End date:" & parseHist2_("OM-11821") & "<br>" )
response.write ("OM-11693;Start date:" & parseHist_("OM-11693") & ";" & "End date:" & parseHist2_("OM-11693") & "<br>" )
response.write ("OM-11516;Start date:" & parseHist_("OM-11516") & ";" & "End date:" & parseHist2_("OM-11516") & "<br>" )
response.write ("OM-11329;Start date:" & parseHist_("OM-11329") & ";" & "End date:" & parseHist2_("OM-11329") & "<br>" )
response.write ("OM-11327;Start date:" & parseHist_("OM-11327") & ";" & "End date:" & parseHist2_("OM-11327") & "<br>" )
response.write ("OM-11324;Start date:" & parseHist_("OM-11324") & ";" & "End date:" & parseHist2_("OM-11324") & "<br>" )
response.write ("OM-11323;Start date:" & parseHist_("OM-11323") & ";" & "End date:" & parseHist2_("OM-11323") & "<br>" )
response.write ("OM-11286;Start date:" & parseHist_("OM-11286") & ";" & "End date:" & parseHist2_("OM-11286") & "<br>" )
response.write ("OM-11280;Start date:" & parseHist_("OM-11280") & ";" & "End date:" & parseHist2_("OM-11280") & "<br>" )
response.write ("OM-11233;Start date:" & parseHist_("OM-11233") & ";" & "End date:" & parseHist2_("OM-11233") & "<br>" )
response.write ("OM-11198;Start date:" & parseHist_("OM-11198") & ";" & "End date:" & parseHist2_("OM-11198") & "<br>" )
response.write ("OM-11171;Start date:" & parseHist_("OM-11171") & ";" & "End date:" & parseHist2_("OM-11171") & "<br>" )
response.write ("OM-11158;Start date:" & parseHist_("OM-11158") & ";" & "End date:" & parseHist2_("OM-11158") & "<br>" )
response.write ("OM-10774;Start date:" & parseHist_("OM-10774") & ";" & "End date:" & parseHist2_("OM-10774") & "<br>" )
response.write ("OM-10773;Start date:" & parseHist_("OM-10773") & ";" & "End date:" & parseHist2_("OM-10773") & "<br>" )
response.write ("OM-10695;Start date:" & parseHist_("OM-10695") & ";" & "End date:" & parseHist2_("OM-10695") & "<br>" )
response.write ("OM-10694;Start date:" & parseHist_("OM-10694") & ";" & "End date:" & parseHist2_("OM-10694") & "<br>" )
response.write ("OM-10681;Start date:" & parseHist_("OM-10681") & ";" & "End date:" & parseHist2_("OM-10681") & "<br>" )
response.write ("OM-10675;Start date:" & parseHist_("OM-10675") & ";" & "End date:" & parseHist2_("OM-10675") & "<br>" )
response.write ("OM-10620;Start date:" & parseHist_("OM-10620") & ";" & "End date:" & parseHist2_("OM-10620") & "<br>" )
response.write ("OM-10619;Start date:" & parseHist_("OM-10619") & ";" & "End date:" & parseHist2_("OM-10619") & "<br>" )
response.write ("OM-10458;Start date:" & parseHist_("OM-10458") & ";" & "End date:" & parseHist2_("OM-10458") & "<br>" )
response.write ("OM-10450;Start date:" & parseHist_("OM-10450") & ";" & "End date:" & parseHist2_("OM-10450") & "<br>" )
response.write ("OM-10443;Start date:" & parseHist_("OM-10443") & ";" & "End date:" & parseHist2_("OM-10443") & "<br>" )
response.write ("OM-10441;Start date:" & parseHist_("OM-10441") & ";" & "End date:" & parseHist2_("OM-10441") & "<br>" )
response.write ("OM-10418;Start date:" & parseHist_("OM-10418") & ";" & "End date:" & parseHist2_("OM-10418") & "<br>" )
response.write ("OM-10412;Start date:" & parseHist_("OM-10412") & ";" & "End date:" & parseHist2_("OM-10412") & "<br>" )
response.write ("OM-10411;Start date:" & parseHist_("OM-10411") & ";" & "End date:" & parseHist2_("OM-10411") & "<br>" )
response.write ("OM-10408;Start date:" & parseHist_("OM-10408") & ";" & "End date:" & parseHist2_("OM-10408") & "<br>" )
response.write ("OM-10407;Start date:" & parseHist_("OM-10407") & ";" & "End date:" & parseHist2_("OM-10407") & "<br>" )
response.write ("OM-10406;Start date:" & parseHist_("OM-10406") & ";" & "End date:" & parseHist2_("OM-10406") & "<br>" )
response.write ("OM-10405;Start date:" & parseHist_("OM-10405") & ";" & "End date:" & parseHist2_("OM-10405") & "<br>" )
response.write ("OM-10371;Start date:" & parseHist_("OM-10371") & ";" & "End date:" & parseHist2_("OM-10371") & "<br>" )
response.write ("OM-10370;Start date:" & parseHist_("OM-10370") & ";" & "End date:" & parseHist2_("OM-10370") & "<br>" )
response.write ("OM-10368;Start date:" & parseHist_("OM-10368") & ";" & "End date:" & parseHist2_("OM-10368") & "<br>" )
response.write ("OM-10335;Start date:" & parseHist_("OM-10335") & ";" & "End date:" & parseHist2_("OM-10335") & "<br>" )
response.write ("OM-10325;Start date:" & parseHist_("OM-10325") & ";" & "End date:" & parseHist2_("OM-10325") & "<br>" )
response.write ("OM-10317;Start date:" & parseHist_("OM-10317") & ";" & "End date:" & parseHist2_("OM-10317") & "<br>" )
response.write ("OM-10315;Start date:" & parseHist_("OM-10315") & ";" & "End date:" & parseHist2_("OM-10315") & "<br>" )
response.write ("OM-10303;Start date:" & parseHist_("OM-10303") & ";" & "End date:" & parseHist2_("OM-10303") & "<br>" )
response.write ("OM-10302;Start date:" & parseHist_("OM-10302") & ";" & "End date:" & parseHist2_("OM-10302") & "<br>" )
response.write ("OM-10301;Start date:" & parseHist_("OM-10301") & ";" & "End date:" & parseHist2_("OM-10301") & "<br>" )
response.write ("OM-10294;Start date:" & parseHist_("OM-10294") & ";" & "End date:" & parseHist2_("OM-10294") & "<br>" )
response.write ("OM-10269;Start date:" & parseHist_("OM-10269") & ";" & "End date:" & parseHist2_("OM-10269") & "<br>" )
response.write ("OM-10262;Start date:" & parseHist_("OM-10262") & ";" & "End date:" & parseHist2_("OM-10262") & "<br>" )
response.write ("OM-10249;Start date:" & parseHist_("OM-10249") & ";" & "End date:" & parseHist2_("OM-10249") & "<br>" )
response.write ("OM-10247;Start date:" & parseHist_("OM-10247") & ";" & "End date:" & parseHist2_("OM-10247") & "<br>" )
response.write ("OM-10246;Start date:" & parseHist_("OM-10246") & ";" & "End date:" & parseHist2_("OM-10246") & "<br>" )
response.write ("OM-10245;Start date:" & parseHist_("OM-10245") & ";" & "End date:" & parseHist2_("OM-10245") & "<br>" )
response.write ("OM-10243;Start date:" & parseHist_("OM-10243") & ";" & "End date:" & parseHist2_("OM-10243") & "<br>" )
response.write ("OM-10234;Start date:" & parseHist_("OM-10234") & ";" & "End date:" & parseHist2_("OM-10234") & "<br>" )
response.write ("OM-10216;Start date:" & parseHist_("OM-10216") & ";" & "End date:" & parseHist2_("OM-10216") & "<br>" )
response.write ("OM-10202;Start date:" & parseHist_("OM-10202") & ";" & "End date:" & parseHist2_("OM-10202") & "<br>" )
response.write ("OM-10198;Start date:" & parseHist_("OM-10198") & ";" & "End date:" & parseHist2_("OM-10198") & "<br>" )
response.write ("OM-10196;Start date:" & parseHist_("OM-10196") & ";" & "End date:" & parseHist2_("OM-10196") & "<br>" )
response.write ("OM-10194;Start date:" & parseHist_("OM-10194") & ";" & "End date:" & parseHist2_("OM-10194") & "<br>" )
response.write ("OM-10179;Start date:" & parseHist_("OM-10179") & ";" & "End date:" & parseHist2_("OM-10179") & "<br>" )
response.write ("OM-10178;Start date:" & parseHist_("OM-10178") & ";" & "End date:" & parseHist2_("OM-10178") & "<br>" )
response.write ("OM-10177;Start date:" & parseHist_("OM-10177") & ";" & "End date:" & parseHist2_("OM-10177") & "<br>" )
response.write ("OM-10173;Start date:" & parseHist_("OM-10173") & ";" & "End date:" & parseHist2_("OM-10173") & "<br>" )
response.write ("OM-10171;Start date:" & parseHist_("OM-10171") & ";" & "End date:" & parseHist2_("OM-10171") & "<br>" )
response.write ("OM-10169;Start date:" & parseHist_("OM-10169") & ";" & "End date:" & parseHist2_("OM-10169") & "<br>" )
response.write ("OM-10168;Start date:" & parseHist_("OM-10168") & ";" & "End date:" & parseHist2_("OM-10168") & "<br>" )
response.write ("OM-10167;Start date:" & parseHist_("OM-10167") & ";" & "End date:" & parseHist2_("OM-10167") & "<br>" )
response.write ("OM-10166;Start date:" & parseHist_("OM-10166") & ";" & "End date:" & parseHist2_("OM-10166") & "<br>" )
response.write ("OM-10165;Start date:" & parseHist_("OM-10165") & ";" & "End date:" & parseHist2_("OM-10165") & "<br>" )
response.write ("OM-10163;Start date:" & parseHist_("OM-10163") & ";" & "End date:" & parseHist2_("OM-10163") & "<br>" )
response.write ("OM-10162;Start date:" & parseHist_("OM-10162") & ";" & "End date:" & parseHist2_("OM-10162") & "<br>" )
response.write ("OM-10161;Start date:" & parseHist_("OM-10161") & ";" & "End date:" & parseHist2_("OM-10161") & "<br>" )
response.write ("OM-10147;Start date:" & parseHist_("OM-10147") & ";" & "End date:" & parseHist2_("OM-10147") & "<br>" )
response.write ("OM-10140;Start date:" & parseHist_("OM-10140") & ";" & "End date:" & parseHist2_("OM-10140") & "<br>" )
response.write ("OM-10138;Start date:" & parseHist_("OM-10138") & ";" & "End date:" & parseHist2_("OM-10138") & "<br>" )
response.write ("OM-10127;Start date:" & parseHist_("OM-10127") & ";" & "End date:" & parseHist2_("OM-10127") & "<br>" )
response.write ("OM-10126;Start date:" & parseHist_("OM-10126") & ";" & "End date:" & parseHist2_("OM-10126") & "<br>" )
response.write ("OM-10125;Start date:" & parseHist_("OM-10125") & ";" & "End date:" & parseHist2_("OM-10125") & "<br>" )
response.write ("OM-10107;Start date:" & parseHist_("OM-10107") & ";" & "End date:" & parseHist2_("OM-10107") & "<br>" )
response.write ("OM-10103;Start date:" & parseHist_("OM-10103") & ";" & "End date:" & parseHist2_("OM-10103") & "<br>" )
response.write ("OM-10047;Start date:" & parseHist_("OM-10047") & ";" & "End date:" & parseHist2_("OM-10047") & "<br>" )
response.write ("OM-9948;Start date:" & parseHist_("OM-9948") & ";" & "End date:" & parseHist2_("OM-9948") & "<br>" )
response.write ("OM-9934;Start date:" & parseHist_("OM-9934") & ";" & "End date:" & parseHist2_("OM-9934") & "<br>" )
response.write ("OM-9933;Start date:" & parseHist_("OM-9933") & ";" & "End date:" & parseHist2_("OM-9933") & "<br>" )
response.write ("OM-9770;Start date:" & parseHist_("OM-9770") & ";" & "End date:" & parseHist2_("OM-9770") & "<br>" )
response.write ("OM-9727;Start date:" & parseHist_("OM-9727") & ";" & "End date:" & parseHist2_("OM-9727") & "<br>" )
response.write ("OM-6560;Start date:" & parseHist_("OM-6560") & ";" & "End date:" & parseHist2_("OM-6560") & "<br>" )
response.write ("OM-6527;Start date:" & parseHist_("OM-6527") & ";" & "End date:" & parseHist2_("OM-6527") & "<br>" )
response.write ("OM-6512;Start date:" & parseHist_("OM-6512") & ";" & "End date:" & parseHist2_("OM-6512") & "<br>" )
response.write ("OM-6385;Start date:" & parseHist_("OM-6385") & ";" & "End date:" & parseHist2_("OM-6385") & "<br>" )
response.write ("OM-6372;Start date:" & parseHist_("OM-6372") & ";" & "End date:" & parseHist2_("OM-6372") & "<br>" )
response.write ("OM-6326;Start date:" & parseHist_("OM-6326") & ";" & "End date:" & parseHist2_("OM-6326") & "<br>" )
response.write ("OM-6245;Start date:" & parseHist_("OM-6245") & ";" & "End date:" & parseHist2_("OM-6245") & "<br>" )
response.write ("OM-6243;Start date:" & parseHist_("OM-6243") & ";" & "End date:" & parseHist2_("OM-6243") & "<br>" )
response.write ("OM-6225;Start date:" & parseHist_("OM-6225") & ";" & "End date:" & parseHist2_("OM-6225") & "<br>" )
response.write ("OM-6132;Start date:" & parseHist_("OM-6132") & ";" & "End date:" & parseHist2_("OM-6132") & "<br>" )
response.write ("OM-6129;Start date:" & parseHist_("OM-6129") & ";" & "End date:" & parseHist2_("OM-6129") & "<br>" )
response.write ("OM-6110;Start date:" & parseHist_("OM-6110") & ";" & "End date:" & parseHist2_("OM-6110") & "<br>" )
response.write ("OM-6109;Start date:" & parseHist_("OM-6109") & ";" & "End date:" & parseHist2_("OM-6109") & "<br>" )
response.write ("OM-6107;Start date:" & parseHist_("OM-6107") & ";" & "End date:" & parseHist2_("OM-6107") & "<br>" )
response.write ("OM-6106;Start date:" & parseHist_("OM-6106") & ";" & "End date:" & parseHist2_("OM-6106") & "<br>" )
response.write ("OM-6105;Start date:" & parseHist_("OM-6105") & ";" & "End date:" & parseHist2_("OM-6105") & "<br>" )
response.write ("OM-6098;Start date:" & parseHist_("OM-6098") & ";" & "End date:" & parseHist2_("OM-6098") & "<br>" )
response.write ("OM-6096;Start date:" & parseHist_("OM-6096") & ";" & "End date:" & parseHist2_("OM-6096") & "<br>" )
response.write ("OM-6033;Start date:" & parseHist_("OM-6033") & ";" & "End date:" & parseHist2_("OM-6033") & "<br>" )
response.write ("OM-5914;Start date:" & parseHist_("OM-5914") & ";" & "End date:" & parseHist2_("OM-5914") & "<br>" )
response.write ("OM-5756;Start date:" & parseHist_("OM-5756") & ";" & "End date:" & parseHist2_("OM-5756") & "<br>" )
response.write ("OM-5539;Start date:" & parseHist_("OM-5539") & ";" & "End date:" & parseHist2_("OM-5539") & "<br>" )
response.write ("OM-5249;Start date:" & parseHist_("OM-5249") & ";" & "End date:" & parseHist2_("OM-5249") & "<br>" )
response.write ("OM-5185;Start date:" & parseHist_("OM-5185") & ";" & "End date:" & parseHist2_("OM-5185") & "<br>" )
response.write ("OM-5180;Start date:" & parseHist_("OM-5180") & ";" & "End date:" & parseHist2_("OM-5180") & "<br>" )
response.write ("OM-5179;Start date:" & parseHist_("OM-5179") & ";" & "End date:" & parseHist2_("OM-5179") & "<br>" )
response.write ("OM-5178;Start date:" & parseHist_("OM-5178") & ";" & "End date:" & parseHist2_("OM-5178") & "<br>" )
response.write ("OM-5177;Start date:" & parseHist_("OM-5177") & ";" & "End date:" & parseHist2_("OM-5177") & "<br>" )
response.write ("OM-5173;Start date:" & parseHist_("OM-5173") & ";" & "End date:" & parseHist2_("OM-5173") & "<br>" )
response.write ("OM-5135;Start date:" & parseHist_("OM-5135") & ";" & "End date:" & parseHist2_("OM-5135") & "<br>" )
response.write ("OM-5108;Start date:" & parseHist_("OM-5108") & ";" & "End date:" & parseHist2_("OM-5108") & "<br>" )
response.write ("OM-5107;Start date:" & parseHist_("OM-5107") & ";" & "End date:" & parseHist2_("OM-5107") & "<br>" )
response.write ("OM-5106;Start date:" & parseHist_("OM-5106") & ";" & "End date:" & parseHist2_("OM-5106") & "<br>" )
response.write ("OM-5069;Start date:" & parseHist_("OM-5069") & ";" & "End date:" & parseHist2_("OM-5069") & "<br>" )
response.write ("OM-5060;Start date:" & parseHist_("OM-5060") & ";" & "End date:" & parseHist2_("OM-5060") & "<br>" )
response.write ("OM-5019;Start date:" & parseHist_("OM-5019") & ";" & "End date:" & parseHist2_("OM-5019") & "<br>" )
response.write ("OM-5018;Start date:" & parseHist_("OM-5018") & ";" & "End date:" & parseHist2_("OM-5018") & "<br>" )
response.write ("OM-4990;Start date:" & parseHist_("OM-4990") & ";" & "End date:" & parseHist2_("OM-4990") & "<br>" )
response.write ("OM-4954;Start date:" & parseHist_("OM-4954") & ";" & "End date:" & parseHist2_("OM-4954") & "<br>" )
response.write ("OM-4939;Start date:" & parseHist_("OM-4939") & ";" & "End date:" & parseHist2_("OM-4939") & "<br>" )
response.write ("OM-4875;Start date:" & parseHist_("OM-4875") & ";" & "End date:" & parseHist2_("OM-4875") & "<br>" )
response.write ("OM-4851;Start date:" & parseHist_("OM-4851") & ";" & "End date:" & parseHist2_("OM-4851") & "<br>" )
response.write ("OM-4850;Start date:" & parseHist_("OM-4850") & ";" & "End date:" & parseHist2_("OM-4850") & "<br>" )
response.write ("OM-4832;Start date:" & parseHist_("OM-4832") & ";" & "End date:" & parseHist2_("OM-4832") & "<br>" )
response.write ("OM-4796;Start date:" & parseHist_("OM-4796") & ";" & "End date:" & parseHist2_("OM-4796") & "<br>" )
response.write ("OM-995;Start date:" & parseHist_("OM-995") & ";" & "End date:" & parseHist2_("OM-995") & "<br>" )
response.write ("OM-937;Start date:" & parseHist_("OM-937") & ";" & "End date:" & parseHist2_("OM-937") & "<br>" )
response.write ("OM-932;Start date:" & parseHist_("OM-932") & ";" & "End date:" & parseHist2_("OM-932") & "<br>" )
response.write ("OM-913;Start date:" & parseHist_("OM-913") & ";" & "End date:" & parseHist2_("OM-913") & "<br>" )
response.write ("OM-833;Start date:" & parseHist_("OM-833") & ";" & "End date:" & parseHist2_("OM-833") & "<br>" )
response.write ("OM-817;Start date:" & parseHist_("OM-817") & ";" & "End date:" & parseHist2_("OM-817") & "<br>" )
response.write ("OM-653;Start date:" & parseHist_("OM-653") & ";" & "End date:" & parseHist2_("OM-653") & "<br>" )
response.write ("OM-647;Start date:" & parseHist_("OM-647") & ";" & "End date:" & parseHist2_("OM-647") & "<br>" )
response.write ("OM-645;Start date:" & parseHist_("OM-645") & ";" & "End date:" & parseHist2_("OM-645") & "<br>" )
response.write ("OM-644;Start date:" & parseHist_("OM-644") & ";" & "End date:" & parseHist2_("OM-644") & "<br>" )
response.write ("OM-643;Start date:" & parseHist_("OM-643") & ";" & "End date:" & parseHist2_("OM-643") & "<br>" )
response.write ("OM-620;Start date:" & parseHist_("OM-620") & ";" & "End date:" & parseHist2_("OM-620") & "<br>" )




Server.ScriptTimeout = 60*60
Response.ContentType = "text/html"
Response.CharSet = "UTF-8"
'response.write (xml)
'response.Write (now)
end sub

function parseHist_(jira_key)
dim url as string
Dim webClient As New System.Net.WebClient
Dim result As String 

url = "https://jira.tomtomgroup.com/rest/api/2/issue/" & jira_key & "?expand=changelog"
webClient.Headers.Add("Content-Type", "application/json; charset=utf-8")
'webClient.Headers.Add("Authorization", "Basic " & "c3ZjX3N0c19qaXJhX3VzZXI6RjIzcnQjNDV0eSM0")
webClient.Headers.Add("Authorization", "Bearer " & "NjEwNTg3NTc5NjU0Ot6RDMksA3OJYGutuCZaMVZs6e1i")
webClient.Headers.Add("X-Atlassian-Token", "nocheck")
result = webClient.DownloadString(url)

parseHist_ = ""
Dim MySerializer As JavaScriptSerializer = New JavaScriptSerializer()
MySerializer.MaxJsonLength = 86753090
Dim parentJson As Dictionary(Of String, Object) = MySerializer.Deserialize(Of Dictionary(Of String, Object))(result)
dim bln as boolean
dim created as string

For Each pair In parentJson
	if (pair.Key.tostring = "changelog" ) then 'changelog is a dictionary [], loop through the pairs directly (key+value)
		for each field in pair.value 'dictionary loop
			if (field.key.tostring = "histories") then 'histories is an array {}, loop through the items, within the items loop through as a dictionary
				for each histitem in field.value 'array loop
				for each p in histitem 'dictionary loop
					if (p.Key.tostring = "created") then
						created = toYYYYMMDD(p.Value.tostring)
					end if
					if (p.key.tostring = "items") then 'items is an array, loop through the items, within the items loop through as a dictionary
						for each itemitem in p.value 'array loop
						bln = false 'use a toggle to start capturing status info
						for each p_item in itemitem 'dictionary loop							
							if (p_item.Value is nothing) then
							else
								if (p_item.key.tostring = "field" and p_item.value.tostring = "Start date") then
									bln = true
									parseHist_ = parseHist_ & created & "||"
								end if
								if (p_item.key.tostring = "from" and bln) then
									parseHist_ = parseHist_ & "from:" & p_item.value.tostring & "||"	
								end if
								if (p_item.key.tostring = "to" and bln) then
									parseHist_ = parseHist_ & "to:" &p_item.value.tostring & "@@"	
								end if
							end if
					'		
						next
						next
					end if
				next 'dictionary loop
				next 'array loop
			end if
		next
	end if
next
end function

function parseHist2_(jira_key)
dim url as string
Dim webClient As New System.Net.WebClient
Dim result As String 

url = "https://jira.tomtomgroup.com/rest/api/2/issue/" & jira_key & "?expand=changelog"
webClient.Headers.Add("Content-Type", "application/json; charset=utf-8")
'webClient.Headers.Add("Authorization", "Basic " & "c3ZjX3N0c19qaXJhX3VzZXI6RjIzcnQjNDV0eSM0")
webClient.Headers.Add("Authorization", "Bearer " & "NjEwNTg3NTc5NjU0Ot6RDMksA3OJYGutuCZaMVZs6e1i")
webClient.Headers.Add("X-Atlassian-Token", "nocheck")
result = webClient.DownloadString(url)

parseHist2_ = ""
Dim MySerializer As JavaScriptSerializer = New JavaScriptSerializer()
MySerializer.MaxJsonLength = 86753090
Dim parentJson As Dictionary(Of String, Object) = MySerializer.Deserialize(Of Dictionary(Of String, Object))(result)
dim bln as boolean
dim created as string

For Each pair In parentJson
	if (pair.Key.tostring = "changelog" ) then 'changelog is a dictionary [], loop through the pairs directly (key+value)
		for each field in pair.value 'dictionary loop
			if (field.key.tostring = "histories") then 'histories is an array {}, loop through the items, within the items loop through as a dictionary
				for each histitem in field.value 'array loop
				for each p in histitem 'dictionary loop
					if (p.Key.tostring = "created") then
						created = toYYYYMMDD(p.Value.tostring)
					end if
					if (p.key.tostring = "items") then 'items is an array, loop through the items, within the items loop through as a dictionary
						for each itemitem in p.value 'array loop
						bln = false 'use a toggle to start capturing status info
						for each p_item in itemitem 'dictionary loop							
							if (p_item.Value is nothing) then
							else
								if (p_item.key.tostring = "field" and p_item.value.tostring = "End date") then
									bln = true
									parseHist2_ = parseHist2_ & created & "||"
								end if
								if (p_item.key.tostring = "from" and bln) then
									parseHist2_ = parseHist2_ & "from:" &p_item.value.tostring & "||"	
								end if
								if (p_item.key.tostring = "to" and bln) then
									parseHist2_ = parseHist2_ & "to:" &p_item.value.tostring & "@@"	
								end if
							end if
					'		
						next
						next
					end if
				next 'dictionary loop
				next 'array loop
			end if
		next
	end if
next
end function

function toYYYYMMDD(s)
'2018-05-29T12:27:48.000+0000 comes in
'20180529 goes out
toYYYYMMDD = mid(s, 1, 4) & mid(s, 6, 2) & mid(s, 9, 2)
end function
</script>
