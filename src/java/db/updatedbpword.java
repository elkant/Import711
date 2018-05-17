/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

package db;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.DataInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.PrintWriter;
import java.net.URL;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.HttpSession;

/**
 *
 * @author Geofrey Nyabuto
 */
public class updatedbpword extends HttpServlet {

   String host,dbase,user,password; 
  static   String dbsetup;
 String url,dbconnpath;
 String status;
    protected void processRequest(HttpServletRequest request, HttpServletResponse response)
            throws ServletException, IOException, ClassNotFoundException, InstantiationException, IllegalAccessException {
        response.setContentType("text/html;charset=UTF-8");
        
        try {
            PrintWriter out = response.getWriter();
            status="";
            if(request.getParameter("hostname")==null){
       
           host="localhost:3306";
           
       }else{
       
        host=request.getParameter("hostname");
        
       }
        
        
        if(request.getParameter("database")==null){
       dbase="pmtct_art_cohort";
        }
        else{
         dbase=request.getParameter("database");
        }
        if(request.getParameter("user")==null){
        user="root";
        }
        else{
        user=request.getParameter("user");
        }
          if(request.getParameter("password")==null){
          
          password="";
          }else{
        password=request.getParameter("password");
          }
      System.out.println("pass entered is : "+password);
//CREATE A PATH IN THE COMPUTER FOR STORING TEXT FILES
                            
    String allpath = getServletContext().getRealPath("/dbase.txt");
        String mydrive = allpath.substring(0, 1);
        //dbconnpath=mydrive+":\\MNHC_SYSTEM_APHIA_PLUS\\"; 
      dbconnpath=mydrive+":\\HALISI\\DATA\\DO_NOT_DELETE\\_\\_\\_"; 
       
      //create a directory
      
      // new File(dbconnpath).mkdir();
     new File(dbconnpath).mkdirs();
        
        
        

    dbsetup =dbconnpath+"\\dbconnection.txt";
        
    //dbsetup=ctx.getRealPath("/dbase.txt");
        
       
        
try {     
                    FileWriter fw = new FileWriter(dbsetup);
			BufferedWriter bw = new BufferedWriter(fw);
			bw.write(host+"\n"+dbase+"\n"+user+"\n"+password.trim());
			bw.close();
                        
                        
                        
    
   } catch (IOException e) {
    
    
}
   
   
   
//   System.out.println("Number of lines:========="+getLineCount(dbsetup));
   
   getLineCount(dbsetup);
        response.setContentType("text/html;charset=UTF-8");
        String url2 = "jdbc:mysql://"+host+"/"+dbase+"";
//         String url = "jdbc:mysql://localhost:3306/javabase";
Connection connection = null;
status="failed";
try {
    System.out.println("Connecting database...");
     Class.forName("com.mysql.jdbc.Driver").newInstance();
    connection = DriverManager.getConnection(url2, user, password);
    System.out.println("Database connected!");
    status="success";
} catch (SQLException e) {
    status="failed";
    System.out.println("status : "+status+" error "+e); 
}
   System.out.println("status : "+status);
            out.println(status);
           
        }
        catch (FileNotFoundException nf){
            
        }
    }

    // <editor-fold defaultstate="collapsed" desc="HttpServlet methods. Click on the + sign on the left to edit the code.">
    /**
     * Handles the HTTP <code>GET</code> method.
     *
     * @param request servlet request
     * @param response servlet response
     * @throws ServletException if a servlet-specific error occurs
     * @throws IOException if an I/O error occurs
     */
    @Override
    protected void doGet(HttpServletRequest request, HttpServletResponse response)
            throws ServletException, IOException {
       try {
           processRequest(request, response);
       } catch (ClassNotFoundException ex) {
           Logger.getLogger(updatedbpword.class.getName()).log(Level.SEVERE, null, ex);
       } catch (InstantiationException ex) {
           Logger.getLogger(updatedbpword.class.getName()).log(Level.SEVERE, null, ex);
       } catch (IllegalAccessException ex) {
           Logger.getLogger(updatedbpword.class.getName()).log(Level.SEVERE, null, ex);
       }
    }

    /**
     * Handles the HTTP <code>POST</code> method.
     *
     * @param request servlet request
     * @param response servlet response
     * @throws ServletException if a servlet-specific error occurs
     * @throws IOException if an I/O error occurs
     */
    @Override
    protected void doPost(HttpServletRequest request, HttpServletResponse response)
            throws ServletException, IOException {
       try {
           processRequest(request, response);
       } catch (ClassNotFoundException ex) {
           Logger.getLogger(updatedbpword.class.getName()).log(Level.SEVERE, null, ex);
       } catch (InstantiationException ex) {
           Logger.getLogger(updatedbpword.class.getName()).log(Level.SEVERE, null, ex);
       } catch (IllegalAccessException ex) {
           Logger.getLogger(updatedbpword.class.getName()).log(Level.SEVERE, null, ex);
       }
    }

    /**
     * Returns a short description of the servlet.
     *
     * @return a String containing servlet description
     */
    @Override
    public String getServletInfo() {
        return "Short description";
    }// </editor-fold>

    
public  void getLineCount (String filename) throws FileNotFoundException, IOException
    {
        
        //URL url3= getClass().getResource("/db.txt");
File file = new File(dbsetup);
        
 FileInputStream fstream = new FileInputStream(file);
  // Get the object of DataInputStream
  DataInputStream in = new DataInputStream(fstream);
  BufferedReader br = new BufferedReader(new InputStreamReader(in));
  String strLine;
  //Read File Line By Line
  while ((strLine = br.readLine()) != null)   {
  // Print the content on the console
  System.out.println ("=="+strLine);
  }
  //Close the input stream
  in.close();
           
    }
    

 public void createdb(){
        try {
            
          URL location = dbConn.class.getProtectionDomain().getCodeSource().getLocation();
          String  mydrive = location.getFile().substring(1, 2);
          
            String command=mydrive+":/wamp/bin/mysql/mysql5.1.36/bin mysql -u root -p"+password+" "+dbase+"  FILE.sql";
            Runtime.getRuntime().exec(command);
        } catch (IOException ex) {
            Logger.getLogger(updatedbpword.class.getName()).log(Level.SEVERE, null, ex);
        }
 
 
 }
 
 
 
 
}

