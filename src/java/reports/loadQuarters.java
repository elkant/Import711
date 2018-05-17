/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

package reports;

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
 * @author Geofrey Nyabuto
 */
public class loadQuarters extends HttpServlet {
HttpSession session;
String data;
int year,prevYear;
String period;
    protected void processRequest(HttpServletRequest request, HttpServletResponse response)
            throws ServletException, IOException, SQLException {
        response.setContentType("text/html;charset=UTF-8");
        PrintWriter out = response.getWriter();
        try {
          session=request.getSession();
          dbConn conn = new dbConn();
          System.out.println(" here q : "+request.getParameter("year"));
          year=Integer.parseInt(request.getParameter("year"));
          prevYear=year-1;
          
          data="<option value =\"\">Choose Quarter</option>";
          
          String getReports="SELECT id,name FROM quarter";
         conn.rs=conn.st.executeQuery(getReports);
         while(conn.rs.next()){
          String  periodInterval [] = conn.rs.getString(2).split("-");
        
          if(conn.rs.getInt(1)==1){
          period=periodInterval[0]+" ("+prevYear+") - ";
          period+=periodInterval[1]+" ("+prevYear+")";
                  }
          else{
           period=periodInterval[0]+" ("+year+") - "; 
          period+=periodInterval[1]+" "+year+")"; 
          }
          data+="<option value=\""+conn.rs.getInt(1)+"\">"+period+"</option>";
        
         } 
          if(conn.rs!=null) {conn.rs.close();}
        if(conn.st!=null) {conn.st.close();}
        if(conn.connect!=null) {conn.connect.close();}
         out.println(data);
        } finally {
            out.close();
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
    } catch (SQLException ex) {
        Logger.getLogger(loadQuarters.class.getName()).log(Level.SEVERE, null, ex);
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
    } catch (SQLException ex) {
        Logger.getLogger(loadQuarters.class.getName()).log(Level.SEVERE, null, ex);
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

}
