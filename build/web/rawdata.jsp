
<%@page import="db.dbConn"%>
<%@page contentType="text/html" pageEncoding="UTF-8"%>
<!DOCTYPE html>
<html>

<!-- BEGIN HEAD -->
<head>
   <meta charset="utf-8" />
   <title>reports</title>
     <link rel="shortcut icon" href="images/index.JPG"/>
   <meta content="width=device-width, initial-scale=1.0" name="viewport" />
   <meta content="" name="description" />
   <meta content="" name="author" />
   <link href="assets/bootstrap/css/bootstrap.min.css" rel="stylesheet" />
   <link href="assets/css/metro.css" rel="stylesheet" />
   <link href="assets/bootstrap/css/bootstrap-responsive.min.css" rel="stylesheet" />

   <link href="assets/font-awesome/css/font-awesome.css" rel="stylesheet" />
   <link href="assets/css/style.css" rel="stylesheet" />
   <link href="assets/css/style_responsive.css" rel="stylesheet" />
   <link href="assets/css/style_default.css" rel="stylesheet" id="style_color" />


   <link rel="stylesheet" href="assets/data-tables/DT_bootstrap.css" />
  

<link rel="stylesheet" href="select2/css/select2.css">

  
</head>
<!-- END HEAD -->
<!-- BEGIN BODY -->
<body class="fixed-top">
   <!-- BEGIN HEADER -->
   <div class="header navbar navbar-inverse navbar-fixed-top">
      <!-- BEGIN TOP NAVIGATION BAR -->
      <div class="navbar-inner">
         <div class="container-fluid">
            <!-- BEGIN LOGO -->
            <h3 style="text-align:center;font-size: 50px;color:white;padding-bottom:16px ;font-weight: bolder;">MOH710,MOH711,MOH705 Reports</h3><br/>
            
            <!-- END LOGO -->
            <!-- BEGIN RESPONSIVE MENU TOGGLER -->
            <a href="javascript:;" class="btn-navbar collapsed" data-toggle="collapse" data-target=".nav-collapse">
            <img src="assets/img/menu-toggler.png" alt="" />
            </a>          
            <!-- END RESPONSIVE MENU TOGGLER -->            
            <!-- BEGIN TOP NAVIGATION MENU -->              
            <ul class="nav pull-right">
              
               <!-- END NOTIFICATION DROPDOWN -->
               <!-- BEGIN INBOX DROPDOWN -->
             
             
               <!-- END TODO DROPDOWN -->
               <!-- BEGIN USER LOGIN DROPDOWN -->
               
             <!--  
               <li class="dropdown user">
                  <a href="#" class="dropdown-toggle" data-toggle="dropdown">
                 
                  <span class="username" style="background: ;">Welcome</span>
                  <i class="icon-angle-down"></i>
                  </a>
                  <ul class="dropdown-menu">
                     <li><a href="editProfile.jsp"><i class="icon-user"></i> User Profile</a></li>
                   
                     <li class="divider"></li>
                     <li><a href="logout.jsp"><i class="icon-key"></i>Log Out</a></li>
                  </ul>
               </li> -->
               
               
               <!-- END USER LOGIN DROPDOWN -->
            </ul>
            <!-- END TOP NAVIGATION MENU --> 
         </div>
      </div>
      <!-- END TOP NAVIGATION BAR -->
   </div>
   <!-- END HEADER -->
   <!-- BEGIN CONTAINER -->
   <div class="page-container row-fluid">
      <!-- BEGIN SIDEBAR -->
      <div class="page-sidebar nav-collapse collapse">
         <!-- BEGIN SIDEBAR MENU -->         
       
         <!-- END SIDEBAR MENU -->
      </div>
      <!-- END SIDEBAR -->
      <!-- BEGIN PAGE -->  
      <div class="page-content">
         <!-- BEGIN SAMPLE PORTLET CONFIGURATION MODAL FORM-->
         <div id="portlet-config" class="modal hide">
            <div class="modal-header">
               <button data-dismiss="modal" class="close" type="button"></button>
               <h3>portlet Settings</h3>
            </div>
            <div class="modal-body">
               <p>Here will be a configuration form</p>
            </div>
         </div>
         <!-- END SAMPLE PORTLET CONFIGURATION MODAL FORM-->
         <!-- BEGIN PAGE CONTAINER-->
         <div class="container-fluid">
            <!-- BEGIN PAGE HEADER-->   
            <div class="row-fluid">
               <div class="span12">
                  <!-- BEGIN STYLE CUSTOMIZER -->
               
                  <!-- END BEGIN STYLE CUSTOMIZER -->   
                  <h3 class="page-title" style="text-align: center;">
                    
<!--                    Internal System-->
                  </h3>
                  <ul class="breadcrumb">
                     <li>
                       
                           <div class="form-actions">
                             <a href="index.jsp"><label  class="btn green">Go to Import data Page</label></a>
                           </div>
                               
                     </li>
           
                  </ul>
               </div>
            </div>
            <!-- END PAGE HEADER-->
            <!-- BEGIN PAGE CONTENT-->
            <div class="row-fluid">
               <div class="span12">
                  <!-- BEGIN SAMPLE FORM PORTLET-->   
                  <div class="portlet box blue">
                     <div class="portlet-title">
                        <h4><i class="icon-reorder"></i>Generate Facility Level Raw data</h4>
                       
                     </div>
                     <div class="portlet-body form">
                        <!-- BEGIN FORM-->
                        <form action="rawdata" id="formActions" class="form-horizontal">
                          
                         
                           <div class="control-group">
                              <label class="control-label">Reporting Year<font color='red'><b>*</b></font></label>
                              <div class="controls">
                                 <select required data-placeholder="Reporting Year" class="chosen-with-diselect span6" tabindex="-1"  id="year" name="year" style="width: 400px;">
                                    <option value=""></option>                                 
                                   
                                 </select>
                              </div>
                           </div>
                          
                             <div class="control-group" id="reportTime">
                              <label class="control-label">Reporting Period<font color='red'><b>*</b></font></label>
                              <div class="controls">
                                  <select required data-placeholder="Reporting Period" onchange="grouplabel();" class="span6 m-wrap" tabindex="-1"  id="reportDuration" name="reportDuration" style="width: 400px;">
                                    <option value="">Choose reporting period</option>                                 
                                   
                                 </select>
                              </div>
                           </div>
                            
                           <div class="control-group" id="reportSemi">
                              <label class="control-label">Semi-annual<font color='red'><b>*</b></font></label>
                              <div class="controls">
                                 <select data-placeholder="Sem annual" class="span6 " tabindex="-1" multiple title="Note:" data-toggle="popover" data-trigger="hover" data-content="You may select more than one period. Press and hold the Ctrl button as you choose the period of interest  " id="semi_annual" name="semi_annual" style="width: 400px;height:45px;">
                                    <option value="">Choose period</option>                                 
                                   
                                 </select>
                              </div>
                           </div>
                            
                            <div class="control-group" id="reportQuarter">
                              <label class="control-label">Quarter<font color='red'><b>*</b></font></label>
                              <div class="controls">
                                 <select data-placeholder="Reporting Quarter" class="span6 " title="Note:" data-toggle="popover" data-trigger="hover" data-content="You may select multiple quarters. Press and hold the Ctrl button as you choose the quarter of interest  " multiple tabindex="-1"  id="quarter" name="quarter" style="width: 400px;height:100%;">
                                    <option value="">Choose Quarter</option>                                 
                                   
                                 </select>
                              </div>
                           </div>
                            
                            
                            <div class="control-group" id="reportMonth">
                              <label class="control-label">Month<font color='red'><b>*</b></font></label>
                              <div class="controls">
                                 <select data-placeholder="Reporting Month" class="span6 " multiple title="Note:" data-toggle="popover" data-trigger="hover" data-content="You may select multiple months. Press and hold the Ctrl button as you choose the months of interest  " tabindex="-1"  id="month" name="month" style="width: 400px;height:220px;">
                                    <option value="">Choose month</option>                                 
                                   
                                 </select>
                              </div>
                           </div>
                            
                             <div class="control-group">
                              <label class="control-label">Report Type </label>
                              <div class="controls">
                                 <select required placeholder="County" onchange="loadsubcounty();"  class="span6 m-wrap" tabindex="-1"  id="reportType" name="reportType" style="width: 400px;">
                                    <option value="">Choose report type</option>
                                    <option value="1">All Facilities</option>
                                    <option value="2" title="You can either select one facility  or get data for all facilities belonging to a specified subcounty or county" >Specify Facilities</option>
                                 </select>
                              </div>
                           </div>
                            
                            
                              <div class="control-group" id="reportCounty">
                              <label class="control-label">County <font color='red'><b>*</b></font></label>
                              <div class="controls">
                                 <select placeholder="County" onchange="loadsubcounty();"  class="span6 m-wrap" tabindex="-1"  id="county" name="county" style="width: 400px;">
                                    <option value=""></option>
                                 </select>
                              </div>
                           </div>
                            
                            <div class="control-group" id="reportDistrict">
                              <label class="control-label">Sub-County </label>
                              <div class="controls">
                                 <select data-placeholder="Sub-County" onchange="loadfacils();"  class="span6 m-wrap" tabindex="-1"  id="subcounty" name="subcounty" style="width: 400px;">
                                    <option value="">Select County First</option>
                                 </select>
                              </div>
                           </div> 
                            
                            
                              <div class="control-group" id="reportFacility">
                              <label class="control-label">Activity Site<font color='red'><b></b></font></label>
                              <div class="controls">
                                 <select data-placeholder="Facility" onchange="updatefacilsession();" class="span6 m-wrap" tabindex="-1"  id="facility" name="facility" style="width: 400px;">
                                    <option value=""></option>
                                 </select>
                                  <i onclick="clearfacil();" style="width:30px;" class="icon-refresh"></i>
                              </div>                                                                                                                                                                                                    
                           </div>
                            
                           <!--<div class="control-group" id="reportFormats">
                              <label class="control-label">Report Format<font color='red'><b>*</b></font></label>
                              <div class="controls">
                                 <select data-placeholder="Report Format" required="true" onchange="getAction();" class="span6 m-wrap" tabindex="-1"  id="reportFormat" name="reportFormat" style="width: 400px;">
                                    <option value="">Choose report Format</option>
                                    <option value="excel">Excel Report</option>
                                    <option value="pdf">PDF Report</option>
                                 </select>
                              </div>
                           </div> -->
                            
                             <div class="control-group">
                              <label class="control-label">Select Form<font color='red'><b>*</b></font></label>
                              <div class="controls">
                                  <select required data-placeholder="Form" class="span6 m-wrap" onchange="getAction();loadsection();" tabindex="-1" id="form" name="form"  style="width: 400px;">
                                   <option value="">Select form</option>   
<%
dbConn conn = new dbConn();

String getforms="select distinct(form) as form from pivottable where active='1' ";
conn.rs=conn.st.executeQuery(getforms);

while(conn.rs.next()){
out.println("<option value="+conn.rs.getString(1)+">"+conn.rs.getString(1).toUpperCase().replace("NEW", "")+"</option>");
}
%>
                                    
                                 </select>
                              </div>
                             </div>
                            
                                 <div class="control-group" >
                              <label class="control-label" >More options<font color='red'></font></label>
                              <div class="controls">
                                <input type='checkbox' id='advanced' onclick='toggleadvanced();'>
                              </div>
                           </div>
                                 
                                 <div id='advancedoption' style='display:none;'>
                                  <div class="control-group" style='display:none;' >
                              <label class="control-label" >Organizational Unit<font color='red'></font></label>
                              <div class="controls">
                                 <select data-placeholder="Report Format" required="true" title="Note:" data-toggle="popover" data-trigger="hover" data-content=" Specify if you want report grouped by facility, Subcounty or county  "  class="span6 m-wrap" tabindex="-1"  id="orgunit" name="orgunit" style="width: 400px;">
                                    
                                    <option title='All facilities' value="facility">Facility</option>
                                    <option title='Subcounty' value="subcounty">Sub county</option>
                                    <option title='county' value="county">County</option>
                                 </select>
                              </div>
                           </div>
                                 
                         
                                 <div class="control-group" >
                              <label class="control-label" id="periodlabel">Group by period?<font color='red'><b>*</b></font></label>
                              <div class="controls">
                                 <select data-placeholder="Report Format" required="true" title="Note:" data-toggle="popover" data-trigger="hover" data-content="Selecting yes means your report will be grouped by the period you will select. eg For a monthly report, data will be grouped by each month"  class="span6 m-wrap" tabindex="-1"  id="groupby" name="groupby" style="width: 400px;">
                                    
                                    <option value="Yes">Yes</option>
                                    <option value="No">No</option>
                                 </select>
                              </div>
                           </div> 
                                 
                           
                                  <div class="control-group" >
                              <label class="control-label" >Indicators List<font color='red'><b>*</b></font></label>
                              <div class="controls">
                                 <select name="indicators" data-placeholder="Report Format" required="true" title="Note:" data-toggle="popover" data-trigger="hover" data-content=" Specify if you want report for all indicators "  class="span6 m-wrap" tabindex="-1"  id="indicators"  style="width: 400px;">
                                    
                                    <option title='All active indicators for the selected form' value="all">All Indicators</option>
                                    <option title='specific indicators that are key for each form. The list of indicators is defined at the backend' value="special">Special Indicators</option>
                                 </select>
                              </div>
                           </div>
                                     
                                     
                              <div class="control-group" >
                              <label class="control-label" >Sections</label>
                              <div class="controls">
                                 <select name="sections" multiple data-placeholder="Report Format" onchange='loadsubsection();'  title="Note:" data-toggle="popover" data-trigger="hover" data-content=" Specify if you want report for all sections "  class="span6 m-wrap" tabindex="-1"  id="sections"  style="width: 400px;">
                                   <option title='All active indicators for the selected form' value="all">All Sections</option>
                                </select>
                              </div>
                           </div>        
                                     
                            <div class="control-group" >
                              <label class="control-label"  >Sub Sections</label>
                              <div class="controls">
                                 <select multiple name="subsection" data-placeholder="Report Format"  title="Note:" data-toggle="popover" data-trigger="hover" data-content=" Specify if you want report for all sub sections "  class="span6 m-wrap" tabindex="-1"  id="subsection"  style="width: 400px;">
                                   <option title='All active indicators for the selected form' value="all">All sub Sections</option>
                                </select>
                              </div>
                           </div> 
                                 
                             </div>    
                                 
                           <table style="width: 100%;">
                           <tr><td class="col-xs-2">
                           <div class="form-actions">
                             <!--<a href="index.jsp"><label  class="btn green">Go to Import data Page</label></a>-->
                           </div>
                                   </td>
                                   
                                   <td class="col-xs-10">
                           <div class="form-actions">
                             
                         
                              
 <button type="submit" class="btn blue">Generate report</button>

                         
                           </div>
                                   </td>
                            </tr> 
                         </table>
                        </form>
                        <!-- END FORM-->           
                     </div>
                  </div>
                  <!-- END SAMPLE FORM PORTLET-->
               </div>
            </div>
       
          
         
          
           
         
          
            <!-- END PAGE CONTENT-->         
         </div>
         <!-- END PAGE CONTAINER-->
      </div>
      <!-- END PAGE -->  
   </div>
   <!-- END CONTAINER -->
   <!-- BEGIN FOOTER -->
   <div class="footer">
      
      <div class="span pull-right">
         <span class="go-top"><i class="icon-angle-up"></i></span>
      </div>
   </div>
   <!-- END FOOTER -->
   <!-- BEGIN JAVASCRIPTS -->    
   <!-- Load javascripts at bottom, this will reduce page load time -->
   <script src="assets/js/jquery-1.8.3.min.js"></script>    
  
   <script src="assets/bootstrap/js/bootstrap.min.js"></script>   

   <script src="assets/js/jquery.cookie.js"></script>
   <!-- ie8 fixes -->
   <!--[if lt IE 9]>
   <script src="assets/js/excanvas.js"></script>
   <script src="assets/js/respond.js"></script>
   <![endif]-->
 
   <script src="assets/js/app.js"></script>  
   <script src="select2/js/select2.js"></script>
  
   
   <script>
     function getAction(){
      var reportFormat="",form="",formActions=""; 

    var formActions=$("#formActions").val();
    var reportduration=$("#reportDuration").val();
    var form=$("#form").val();
//    alert("format : "+reportFormat+" form : "+form);
    
    if(reportduration==="4"){
      document.getElementById("formActions").action ="rawdata";   
    }
    else{
   
 document.getElementById("formActions").action ="rawdata";



     }  
       
    }
       
      jQuery(document).ready(function() {       
         // initiate layout and plugins
        $('#facility').select2(); 

       // FormComponents.init();
         $("#reportTime").hide();
         $("#reportSemi").hide();
         $("#reportQuarter").hide();
         $("#reportMonth").hide();
         $("#reportCounty").hide();
  $("#reportDistrict").hide();
  $("#reportFacility").hide(); 
      //load all the facilities first to enable one to filter by county
                 $.ajax({
url:'loadFacilities',
type:'post',
dataType:'html',
success:function (data){
       $("#facility").html(data);
         if(document.getElementById("facility").value!==''){
      updatefacilsession();
      
      $('#facility').select2();  
      }  
     
         // $("#facility").chosen();
       
       
}


});     
         
$.ajax({
    url:'loadYear',
    type:'post',
    dataType:'html',
    success:function (data){
        document.getElementById("year").innerHTML=data;
     var year=$("#year").val();
     if(year===""){
         
     }
     else{
    $("#reportTime").show();  
  reportingPeriod();     
     }
    }
    
    
});

$("#reportType").change(function(){
  var report=$("#reportType").val();
  if(report===""){
  $("#reportCounty").hide();
  $("#reportDistrict").hide();
  $("#reportFacility").hide(); 
 
        }
       else if(report==="1"){
   $("#reportCounty").hide();
  $("#reportDistrict").hide();
  $("#reportFacility").hide();
  $("#county").removeAttr("required");
  //$("#facility").removeAttr("required");
       } 
        
    else{
  $("#reportCounty").show();
  $("#reportDistrict").show();
  $("#reportFacility").show(); 
   $("#county").attr("required",true);
  // $("#facility").attr("required",true);
    }    
  
    });
$("#reportDuration").change(function(){
  var report=$("#reportDuration").val();
  if(report===""){
    $("#reportSemi").hide();
    $("#reportQuarter").hide();
    $("#reportMonth").hide();  
                 }
  else if(report==="1"){
   $("#reportSemi").hide();
    $("#reportQuarter").hide();
    $("#reportMonth").hide();
    
   $("#month").removeAttr("required");
   $("#quarter").removeAttr("required");
   $("#semi_annual").removeAttr("required");
//   $("#facility").attr("required",true);
  }
  else if(report==="2"){
  $("#reportSemi").show();
  $("#reportQuarter").hide();
 $("#reportMonth").hide(); 
 
    $("#month").removeAttr("required");
   $("#quarter").removeAttr("required");
   $("#semi_annual").attr("required",true);
  }
  
  else if(report==="3"){
  $("#reportSemi").hide();
  $("#reportQuarter").show();
 $("#reportMonth").hide();
 
  $("#quarter").attr("required",true);
  $("#month").removeAttr("required");
  $("#semi_annual").removeAttr("required");
  
  }
 
  else if(report==="4"){
  $("#reportSemi").hide();
  $("#reportQuarter").hide();
  $("#reportMonth").show(); 
  
   $("#month").attr("required",true);
   $("#quarter").removeAttr("required");
   $("#semi_annual").removeAttr("required");
  }
  
        
});
  
$("#year").change(function(){
  var year=$("#year").val();
  if(year===""){
  $("#reportTime").hide();    
  }
  else{
 $("#reportTime").show();  
  reportingPeriod();
  }
    });
  
});
      
      
      
      function reportingPeriod(){
          
          var yea=$("#year").val();
    $.ajax({
         
url:"loadReportingPeriod?yr="+yea,
type:'post',
dataType:'html',
success:function (data){
 $('#reportDuration').html(data); 
 $('#reportDuration').select2();       
}
    });
        $("#reportTime").show();
        loadSemiAnnual();
        loadQuarters();
        loadmonths();     
      }
      
      function loadSemiAnnual(){
      var year=$("#year").val();
    
       $.ajax({
url:'loadSemiAnnual?year='+year,
type:'post',
dataType:'html',
success:function (data){
 $('#semi_annual').html(data.replace('<option value ="">Choose Semi Annual</option>','')); 
 //$('#semi_annual').select2();       

}
    });
          
      }
     
     function loadQuarters(){
     var year=$("#year").val();
       $.ajax({
url:'loadQuarters?year='+year,
type:'post',
dataType:'html',
success:function (data){
 $('#quarter').html(data.replace('<option value ="">Choose Quarter</option>','')); 
 
 //$('#quarter').select2();       
}
    });
          
      }
      
    
      function loadfacils(){
      var subcounty=document.getElementById("subcounty").value;  
                    $.ajax({
url:'loadFacilities?subcounty='+subcounty,
type:'post',
dataType:'html',
success:function (data){
       $("#facility").html(data);
         if(document.getElementById("facility").value!==''){
      updatefacilsession();
      
     
      }  
      $('#facility').select2();  
         // $("#facility").chosen();
       
       
}


}); 
         
         
        }
      
    
    function updatefacilsession(){
          
        var facil=document.getElementById("facility").value;
        $.ajax({
url:'updatefacilitysession?facil='+facil,
type:'post',
dataType:'html',
success:function (data){      
    
    //  $("#"+col).css({'background-color' : '#CCFFCC'});
     
     //now load the forms
     
     loadfrms();
     
}
             
             });    
          
          
          
      }
      
    
    function loadfrms(){
        
        
        $.ajax({
            url:'loadReportedForms',
            type:'post',
            dataType:'html',
            success:function (data){
//                $("#form").html(data);
                
              //  App.init(); 
              
              //also load county and facility
              loadcounty();
              //loadsubcounty();
            }
            
            
        });
        
    }
    
    
    
     
    function loadcounty(){
        
        
        $.ajax({
            url:'loadCounty',
            type:'post',
            dataType:'html',
            success:function (data){
                $("#county").html(data);
                loadsubcounty();
              //  App.init();   
            }
            
            
        });
        
    }
    
    
       function loadsubcounty(){
        
        var county=document.getElementById("county").value;
        $.ajax({
            url:'loadSubcounty?county='+county,
            type:'post',
            dataType:'html',
            success:function (data){
                $("#subcounty").html(data);
                
              //  App.init();   
            }
            
            
        });
        
    }
    
    
    
      function loadmonths(){
      //alert("");
      var yr=document.getElementById("year").value;
      
              $.ajax({
url:'loadMonth?year='+yr,
type:'post',
dataType:'html',
success:function (data){
    $("#month").html(data.replace("<option value=''>Select Month </option>",""));     
    
       //document.getElementById("month").innerHTML=data;
      // App.init();  
        
}


});  
      
      
      }
      
      //load default facilities
     loadcounty();
      
      
           $(document).ready(function(){
          
          $('[data-toggle="popover"]').popover();
          
      });
      
      
      
      function clearfacil(){
          
          console.log("clear selected facil");
         $("#facility").val([]); 
         
           $('#facility').select2(); 
      }
      
   function grouplabel(){
      var periodname="";
       var periodid=$("#reportDuration").val();
       
       if(periodid==='1'){periodname="Group by Year";}
       else if(periodid==='2'){periodname="Group by Semi-Annual";}
      else  if(periodid==='3'){periodname="Group by Quarter";}
      else  if(periodid==='4'){periodname="Group by Month";}
      else  {periodname="Group by Period";}
       
       $("#periodlabel").html(periodname);
       
   }   
      
    grouplabel();
    
    function toggleadvanced(){
        
       //if($('#advancedoption').is(":checked"))
        if($('#advanced:checkbox:checked').length > 0)
       {
        $("#advancedoption").show();   
           console.log("show");
       }
       else {
           
         $("#advancedoption").hide(); 
          console.log("hide");
       }
        
                             }
    toggleadvanced();
    
   
   //add section and subsectionfilter
   
   
      function loadsection(){
        
        var frm=$("#form").val();
        
        if(frm!==''){
        $.ajax({
            url:'loadSections?form='+frm,
            type:'post',
            dataType:'html',
            success:function (data){
                $("#sections").html(data);
                
              //also load county and facility
              
              //loadsubcounty();
                                    }
            
            
        });  }
        
                           }
   
   
   //load subsections
   
   
   
    function loadsubsection(){
        
        var sectionid=$("#sections").val();
        if(sectionid!==''){
        $.ajax({
            url:'loadSubsections?sectionid='+sectionid,
            type:'post',
            dataType:'html',
            success:function (data){
                
                $("#subsection").html(data);
                
             
                                    }
            
            
               });
           }
        
                             }
    loadsection();
    loadsubsection();
   </script>
   
   <!-- END JAVASCRIPTS -->   
</body>
<!-- END BODY -->
</html>

