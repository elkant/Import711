/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package loaddata;

import db.dbConn;
import java.io.IOException;
import java.io.PrintWriter;
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
 * @author Elkant
 */
public class loadCounty extends HttpServlet {
HttpSession session=null;
    
    protected void processRequest(HttpServletRequest request, HttpServletResponse response)
            throws ServletException, IOException {
        try {
    response.setContentType("text/html;charset=UTF-8");
    session=request.getSession();
 String sessioncounty="";   
    
    
    if(session.getAttribute("countyid")!=null){
    
    sessioncounty=session.getAttribute("countyid").toString();
    }
    
    
    String county="";
    
    dbConn conn= new dbConn();
county="<option value=\"\">Select County</option>";
String qr="select * from county";
conn.rs=conn.st.executeQuery(qr);
while(conn.rs.next()){
if(sessioncounty.equalsIgnoreCase(conn.rs.getString(1))){

county+="<option selected value=\""+conn.rs.getString(1)+"\">"+conn.rs.getString(2)+"</option>";
}    

else {
county+="<option value=\""+conn.rs.getString(1)+"\">"+conn.rs.getString(2)+"</option>";
}

                      }
    
    PrintWriter out = response.getWriter();
    try {
      out.println(county); 
    } finally {            
        out.close();
           if (conn!=null) try {
               
             
               if(conn.rs!=null){ conn.rs.close();}
               if(conn.st!=null){ conn.st.close();}
              
               
               
           }catch (Exception ignore) {}
    }
}       catch (SQLException ex) {
            Logger.getLogger(loadCounty.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    // <editor-fold defaultstate="collapsed" desc="HttpServlet methods. Click on the + sign on the left to edit the code.">
    /** 
     * Handles the HTTP <code>GET</code> method.
     * @param request servlet request
     * @param response servlet response
     * @throws ServletException if a servlet-specific error occurs
     * @throws IOException if an I/O error occurs
     */
    @Override
    protected void doGet(HttpServletRequest request, HttpServletResponse response)
            throws ServletException, IOException {
        processRequest(request, response);
    }

    /** 
     * Handles the HTTP <code>POST</code> method.
     * @param request servlet request
     * @param response servlet response
     * @throws ServletException if a servlet-specific error occurs
     * @throws IOException if an I/O error occurs
     */
    @Override
    protected void doPost(HttpServletRequest request, HttpServletResponse response)
            throws ServletException, IOException {
        processRequest(request, response);
    }

    /** 
     * Returns a short description of the servlet.
     * @return a String containing servlet description
     */
    @Override
    public String getServletInfo() {
        return "Short description";
    }// </editor-fold>
}
