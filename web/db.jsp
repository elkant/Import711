<%-- 
    
--%>
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <title>IMPORT Database Set Up</title>
  <link rel="stylesheet" href="modal/jquery-ui.css">
  <script src="assets/js/jquery-1.8.3.min.js"></script>
  <script src="assets/jquery-ui/jquery-ui-1.10.1.custom.min.js"></script>
  <link rel="stylesheet" href="assets/jquery-ui/jquery-ui-1.10.1.custom.min.css">
  <link rel="stylesheet" href="css/btn.css">
  <link rel="shortcut icon" href="images/header.jpg"/>
  <script>
  $(document).ready(function(){
       $("#success").hide();
    var dialog, form,
       emailRegex = /^[a-zA-Z0-9.!#$%&'*+\/=?^_`{|}~-]+@[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?(?:\.[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?)*$/,
      name = $( "#user" ),
      password = $( "#password" ),
      allFields = $( [] ).add( name ).add( password ),
      tips = $( ".validateTips" );

 
    function EditConnection(){
     $(".validateTips1").val("here");   
        $(".validateTips1").html("<font color=\"black\"><b>Saving Details...</b><img src=\"images/utube.gif\"></font>"); 
      var valid = true;
      allFields.removeClass( "ui-state-error" );
      if ( valid ) {
        dialog.dialog( "open" );
        var hostname,database,user,password;
        hostname=$("#hostname").val();
        database=$("#database").val();
        user=$("#user").val();
        password=$("#password").val();
//        alert("username : "+username+" password : "+password);
       
       if(hostname==""){
           $(".validateTips").html("<font color=\"red\"><b>Please enter hostname...</b></font>"); 
           
            }
            else if(database==""){
             $(".validateTips").html("<font color=\"red\"><b>Please enter database name...</b></font>");    
                
            } 
            else if(user==""){
             $(".validateTips").html("<font color=\"red\"><b>Please enter database user...</b></font>");    
                
            }
              else{
                  $(".validateTips1").show();
              $.ajax({
        url:"updatedbpword?hostname="+hostname+"&&database="+database+"&&user="+user+"&&password="+password,
        type:"post",
        dataType:"html",
        success:function(data){
            data=$.trim(data);
if(data=="success"){
     var timer=5;
    var interv= setInterval(function(){ $(".validateTips1").html("<font color=\"green\"><b>Correct Connection Credentials. You will be redirected within "+timer+" seconds</b></font>"); timer--; }, 1000);
    $(".validateTips1").fadeOut(8000);
    setTimeout(function(){ clearInterval(interv); $(".validateTips1").val("");timer=5; window.location.href="index.jsp";}, 5000);

}
else{
  var timer=5;   
$(".validateTips1").html("<font color=\"red\"><b>Failed: Wrong Connection credentials.</b></font>");
setInterval(function(){ $(".validateTips2").html("<font color=\"red\"><b>Please try again...</b></font>");$(".validateTips2").fadeOut(5000); $(".validateTips2").show();$(".validateTips2").val(""); }, 7000);
}
        }
    });
       }
    
      }
      return valid;
    }
 
    dialog = $( "#dialog-form" ).dialog({
      autoOpen: false,
//      height: 300,
      width: 550,
      modal: true,
      buttons: {
        "Save": EditConnection
      }, 
   closeOnEscape: false,
   open: function(event, ui) { $(".ui-dialog-titlebar-close", ui.dialog || ui).hide(); }

      
    });
 
    form = dialog.find( "form" ).on( "submit", function( event ) {
      event.preventDefault();
//      addUser();
    });
      dialog.dialog( "open" );  
  });
  </script>
  <style>
      .ui-dialog-titlebar-close {
  visibility: hidden;
}
  </style>
</head>
<body>
 
    <div id="dialog-form" title="Enter Database Connection Credentials...." hidden="true" >
 <!--<div contenteditable=true>-->
 <p class="validateTips1" style="text-align: center;"></p>
 <p class="validateTips2" style="text-align: center;"></p>
 <p class="validateTips3" style="text-align: center;"></p>
 <p class="validateTips" style="text-align: center;"></p>
  <form action="" method="post" >
                    <h4 align="center">Database Configuration</h4>
                    <table  cellpadding="8px" cellspacing="6px">
                        <tr>
                            <td style="text-align: right;">Host name:   <img src="images/blguide.png" title="enter the hostname followed by a ':' then port number e.g. localhost:3306. If there is no port number, just enter the host name alone e.g. localhost"/></td>
                            <td style="text-align: right">

                                <!--  username -->
                                <input id="hostname" type=text required name="hostname" id="hostname" placeholder="e.g localhost:3306" value="localhost:3306" autofocus class="textbox"/></td>

                        </tr>
                        <tr> <td style="text-align: right;">Database : <img src="images/blguide.png" title="enter the database name e.g internal_system"/></td>
                            <td style="text-align: right;"><input id="database"  type=text required name="database" value="halisi" placeholder="e.g halisi"  class="textbox"/></td>
                        </tr>

                        <tr> <td style="text-align: right;">User:   <img src="images/blguide.png" title="enter a database user name name e.g root"/></td>
                            <td style="text-align: right;"><input id="user"  type=text required name="user" value="root" placeholder="e.g root"  class="textbox"/></td>
                        </tr>

                        <tr><td style="text-align: right;">Password:   <img src="images/blguide.png" title="enter the database password"/></td> 
                            <!--password-->
                            <td style="text-align: right;"><input id="password"  type=password  name="password" placeholder="Password" class="textbox"></td>
                        </tr>
                        
                    </table>
                
                </form>
  </div>  
    <div id="success">
        <p style="color: red; ">Error While connecting to the database................</p>
        </div>
</body>
</html>