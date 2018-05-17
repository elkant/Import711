<%-- 
    Document   : loadTBExcel
    Created on : Jul 27, 2015, 2:41:29 PM
    Author     : Maureen
--%>




<%@page import="db.dbConn"%>
<%@page import="java.util.Calendar"%>
<%@page contentType="text/html" pageEncoding="UTF-8"%>
<!DOCTYPE html>
<html>

<!-- BEGIN HEAD -->
<head>
   <meta charset="utf-8" />
   <title>Upload 711 DATA </title>
   <link rel="shortcut icon" href="images/logo.png"/>
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

   <link rel="stylesheet" type="text/css" href="assets/uniform/css/uniform.default.css" />
<link rel="stylesheet" href="select2/css/select2.css">
<link rel="stylesheet" href="css/animate.css">


                
                <style>
                    
                    [data-notify="progressbar"] {
	margin-bottom: 0px;
	position: absolute;
	bottom: 0px;
	left: 0px;
	width: 100%;
	height: 5px;
}
                    
                </style>
                
  
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
            <h3 style="text-align:center;font-size: 50px;color:white;padding-bottom:16px ;font-weight: bolder;">MOH 711, 710 AND 705 monthly data</h3><br/>
            
            <!-- END LOGO -->
            <!-- BEGIN RESPONSIVE MENU TOGGLER -->
            <a href="javascript:;" class="btn-navbar collapsed" data-toggle="collapse" data-target=".nav-collapse">
            <img src="assets/img/menu-toggler.png" alt="" />
            </a>          
                      
            <ul class="nav pull-right">
          
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
                     <li style="width: 900px;">
                        <i class="icon-home"></i>
                        <a href="#" style="margin-left:40%;">Upload excel files.</a> 
                        <!--<span class="icon-angle-right"></span>-->
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
                        <h4><i class="icon-reorder"></i> Import data from Excel files (.xlsx or .xlsm)</h4>
                       
                     </div>
                     <div class="portlet-body form">
                        <!-- BEGIN FORM-->
                        <form action="importdata" method="post" enctype="multipart/form-data" class="form-horizontal" >
                       
                            
<!--                            <div class="control-group">
                              <label class="control-label">Week Start date:<font color='red'><b>*</b></font></label>
                              <div class="controls">
                                  <input required type="text" title="this is the date that the week started" value="<%if (session.getAttribute("weekstart") != null) {out.println(session.getAttribute("weekstart")); }%>" class="form-control input-lg tarehe" name="weekstart" id="weekstart">
                              </div>
                           </div>-->
                            
                            
                             <div class="control-group">
                              <label class="control-label">Year<font color='red'><b>*</b></font></label>                              
                              <div class="controls">
                              <select onchange="" title="Year when data was collected" name="year" id="year" class="form-control col-xs-6" style="" >
                                            <option value='' >Choose year</option>
                                            <%
                                                
                                                Calendar cal= Calendar.getInstance();
                                                int curyear=cal.get(Calendar.YEAR);
                                                
                                            for(int a=2017;a<=curyear;a++){
                                             out.println("<option value='"+a+"'>"+a+"</option>");
                                                %>
                                            
                                            <%
                                            }
                                            
                                            %>
                                            
                                        </select> 
                           </div>
                           </div>

                                            
                                            
                                               <div class="control-group">
                              <label class="control-label">Month<font color='red'><b>*</b></font></label>                              
                              <div class="controls">
                              <select required="true"    name="month" id="month" onchange="" class="form-control" >
                                            <option>Select Month</option>
                                            <option value="01">January</option>
                                            <option value="02">February</option>
                                            <option value="03">March</option>
                                            <option value="04">April</option>
                                            <option value="05">May</option>
                                            <option value="06">June</option>
                                            <option value="07">July</option>
                                            <option value="08">August</option>
                                            <option value="09">September</option>
                                            <option value="10">October</option>
                                            <option value="11">November</option>
                                            <option value="12">December</option>
                                           
                                        </select>
                           </div>
                           </div>
                                            
                             <div class="control-group">
                              <label class="control-label">Excel file<font color='red'><b>*</b></font></label>
                              <div class="controls">
                                  <input required type="file" name="file_name" id="upload" value="" class="textbox" required>  
                              </div>
                           </div>
                          
                           
                        <br><br><br><br>



                         <table style="width: 100%;">
                           <tr><td class="col-xs-6">
                           <div class="form-actions">
                              <button type="submit" class="btn blue">Upload excel Excel.</button>

                         
                           </div>
                                   </td>
                                   
                                   <td class="col-xs-6">
                           <div class="form-actions">
                             
                         
                              <a href="rawdata.jsp"><label  class="btn green">Go to Reports Page</label></a>

                           </div>
                                   </td>
                            </tr> 
                         </table>
                        <div class="form-actions" id="matokeo">
                        <div class="form-actions">
                            
                        </div>
                        
                        <!-- END FORM-->           
                     </div>
                        </form>
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
       <%

              cal = Calendar.getInstance();
                    int year = cal.get(Calendar.YEAR);       
%>
     <% dbConn conn= new dbConn(); %>  
    <div class="span pull-right">
         <span class="go-top"><i class="icon-angle-up"></i></span>
      </div>
   </div>
   <!-- END FOOTER -->
   <!-- BEGIN JAVASCRIPTS -->    
   <!-- Load javascripts at bottom, this will reduce page load time -->
   
<script src="assets/js/jquery-1.8.3.min.js"></script>
   

<script type="text/javascript" src="js/bootstrap-notify.js"></script>


      
         
   <script src="assets/bootstrap/js/bootstrap.min.js"></script>   
   <script type="text/javascript" src="assets/bootstrap-fileupload/bootstrap-fileupload.js"></script>
   <script src="assets/js/jquery.blockui.js"></script>
   <script src="assets/js/jquery.cookie.js"></script>
   <!-- ie8 fixes -->
   <!--[if lt IE 9]>
   <script src="assets/js/excanvas.js"></script>
   <script src="assets/js/respond.js"></script>
   <![endif]-->
   <script type="text/javascript" src="assets/chosen-bootstrap/chosen/chosen.jquery.min.js"></script>
   <script type="text/javascript" src="assets/uniform/jquery.uniform.min.js"></script>
   <script type="text/javascript" src="assets/bootstrap-wysihtml5/wysihtml5-0.3.0.js"></script> 
   <script type="text/javascript" src="assets/bootstrap-wysihtml5/bootstrap-wysihtml5.js"></script>
   <script type="text/javascript" src="assets/jquery-tags-input/jquery.tagsinput.min.js"></script>
   <script type="text/javascript" src="assets/bootstrap-toggle-buttons/static/js/jquery.toggle.buttons.js"></script>
   <script type="text/javascript" src="assets/bootstrap-datepicker/js/bootstrap-datepicker.js"></script>
   <script type="text/javascript" src="assets/clockface/js/clockface.js"></script>
   <script type="text/javascript" src="assets/bootstrap-daterangepicker/date.js"></script>
   <script type="text/javascript" src="assets/bootstrap-daterangepicker/daterangepicker.js"></script> 
   <script type="text/javascript" src="assets/bootstrap-colorpicker/js/bootstrap-colorpicker.js"></script>  
   <script type="text/javascript" src="assets/bootstrap-timepicker/js/bootstrap-timepicker.js"></script>
   <script src="assets/js/app.js"></script>  
   <script src="select2/js/select2.js"></script>
  
     

<script > 
                
</script>

<script>
      
   
  
      
      $(".tarehe").datepicker({
    clearBtn: true
}).on('changeDate', function(ev){
    $(this).datepicker('hide');
});
      
      
   </script>

                  
 <%if (session.getAttribute("uploadedpns") != null) { %>
                                <script type="text/javascript"> 
                    
                    
$("#matokeo").html('<%=session.getAttribute("uploadedpns")%>');
                         
      $.notify(
      {
  message:'<%=session.getAttribute("uploadedpns")%>'},
      {
	icon_type: 'image'
      }, 
      {
	offset: {
		x: 600,
		y: 300
	}
       }
       
            ); 
                    
                </script>
                
                <%
                //session.removeAttribute("uploadedart");
                            }

                        %>


<%if (session.getAttribute("reportingyear") != null) { %>
                <script type="text/javascript">
                    $("#year").val(<%=session.getAttribute("reportingyear")%>);
                    $("#month").val(<%=session.getAttribute("reportingmonth")%>);
                    
                    </script>
                
                <%}%>
   
 <%if (session.getAttribute("resp1") != null) { %>
                                <script type="text/javascript"> 
                    
                    
                    
                         
      $.notify(
      {icon: "images/validated.jpg", 
  message:'<%=session.getAttribute("resp1")%>'},
      {
	icon_type: 'image'
      }, 
      {
	offset: {
		x: 600,
		y: 300
	}
       }
       
            ); 
    
    
     $.notify(
      {icon: "images/validated.jpg", 
  message:'<%=session.getAttribute("resp")%>'},
      {
	icon_type: 'image'
      }, 
      {
	offset: {
		x: 600,
		y: 300
	}
       }
       
            ); 
    
                    
                </script>
                
                <%
                session.removeAttribute("resp1");
                session.removeAttribute("resp");
                            }

 
 
                        %>
     

  
   
   <!-- END JAVASCRIPTS -->   
</body>
<!-- END BODY -->
</html>


