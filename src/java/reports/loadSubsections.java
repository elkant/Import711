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
public class loadSubsections extends HttpServlet {
HttpSession session;
String subsectionQuery,current_subsections;
    protected void processRequest(HttpServletRequest request, HttpServletResponse response)
            throws ServletException, IOException, SQLException {
        response.setContentType("text/html;charset=UTF-8");
        PrintWriter out = response.getWriter();
        try {
              dbConn conn=new dbConn();
              session=request.getSession();
              
            subsectionQuery="";current_subsections="";
            
         if(request.getParameter("sectionid").contains(",")){
             String countids []=request.getParameter("sectionid").split(",");
             
             for(String ids: countids){
             subsectionQuery+=" || sectionid="+ids;   
                 
             }
             System.out.println("before remove : "+subsectionQuery);
             
             subsectionQuery=subsectionQuery.substring(3);
             
             System.out.println("after remove : "+subsectionQuery);
         }
         else{
             subsectionQuery="sectionid="+request.getParameter("sectionid");
         }
         
         System.out.println("final : "+subsectionQuery);
         
        String facilities="Select subsectionid,subsection from pivottable where "+subsectionQuery+" GROUP BY subsectionid";
      System.out.println("query : "+facilities);
          conn.rs=conn.st.executeQuery(facilities);
      while(conn.rs.next()){

current_subsections+="<option value=\""+conn.rs.getString(1)+"\">"+conn.rs.getString(2)+"</option>";
 }
        System.out.println(current_subsections);
            out.println(current_subsections);
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
        Logger.getLogger(loadSubsections.class.getName()).log(Level.SEVERE, null, ex);
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
        Logger.getLogger(loadSubsections.class.getName()).log(Level.SEVERE, null, ex);
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
