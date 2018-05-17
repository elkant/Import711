/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package loaddata;

import db.dbConn;
import java.io.IOException;
import java.io.PrintWriter;
import java.sql.SQLException;
import java.util.Calendar;
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
public class loadYear extends HttpServlet {
HttpSession session;
    protected void processRequest(HttpServletRequest request, HttpServletResponse response)
            throws ServletException, IOException {
        try {
            response.setContentType("text/html;charset=UTF-8");

session = request.getSession();
              Calendar cal = Calendar.getInstance();
                    int year = cal.get(Calendar.YEAR);
         String sessionyear="";   
                   int currentmonth= cal.get(Calendar.MONTH)+1;
if(currentmonth>=10 && currentmonth<=12){

year+=1;
    
                                        } 
                   
            dbConn conn = new dbConn();
if(session.getAttribute("year")!=null){
    
    sessionyear=session.getAttribute("year").toString();
                                      }
            String getyears = "select distinct(year) as year from moh711_new";

            String years = "<option value=''>Select Year</option>";


            conn.rs = conn.st.executeQuery(getyears);

            while (conn.rs.next()) {
                 if(sessionyear.equalsIgnoreCase(conn.rs.getString("year"))){
                  if(conn.rs.getInt("year")<=year){
                years += "<option selected value='" + conn.rs.getString("year") + "'>" + conn.rs.getString("year") + "</option> ";
                }}
                 else{
                if(conn.rs.getInt("year")<=year){
                years += "<option value='" + conn.rs.getString("year") + "'>" + conn.rs.getString("year") + "</option> ";
                }}
                }
            PrintWriter out = response.getWriter();
            try {

                out.println(years);

            } finally {
               if(conn.connect!=null){ conn.connect.close();}
               if(conn.rs!=null){ conn.rs.close();}
               if(conn.st!=null){ conn.st.close();}
                out.close();
            }
        } catch (SQLException ex) {
            Logger.getLogger(loadYear.class.getName()).log(Level.SEVERE, null, ex);
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
        processRequest(request, response);
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
        processRequest(request, response);
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
