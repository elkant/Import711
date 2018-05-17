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
public class loadFacilities extends HttpServlet {

    HttpSession session=null;
    
    protected void processRequest(HttpServletRequest request, HttpServletResponse response)
            throws ServletException, IOException {
        response.setContentType("text/html;charset=UTF-8");
        
        session=request.getSession();
             dbConn conn = new dbConn();
        String facilityonsession="";
        
        String subcounty="";
        
       subcounty=request.getParameter("subcounty"); 
       
       //disttrict is same as subcounty
       
       
        if(session.getAttribute("facilityid")!=null){
        facilityonsession=session.getAttribute("facilityid").toString();
        
        }
        System.out.println("facil on session__"+facilityonsession);
        
        PrintWriter out = response.getWriter();
        try {
       
           
            String facils="<option value=''>Select Site</option>";
            String getfacils="";
            
            if(subcounty!=null){
             getfacils="select SubPartnerId,SubPartnerNom from subpartnera where  DistrictID='"+subcounty+"' order by SubPartnerNom ";
            }
            else {
              getfacils="select SubPartnerId,SubPartnerNom from subpartnera order by SubPartnerNom ";
            
            
            }
            conn.rs=conn.st.executeQuery(getfacils);
            System.out.println(getfacils);
            while(conn.rs.next()){
                //if the current facility on loop is same as the facility on session, then make it selected
                if(facilityonsession.equalsIgnoreCase(conn.rs.getString(1).trim())){
                
            facils+="<option selected value='"+conn.rs.getString(1) +"'> "+conn.rs.getString(2)+" </option>";
                }
                else{
                         facils+="<option value='"+conn.rs.getString(1) +"'> "+conn.rs.getString(2)+" </option>";
                }
               // System.out.println("~~"+conn.rs.getString(2));
            
                                 }
            
             
            out.println(facils);
        } catch (SQLException ex) {
            Logger.getLogger(loadFacilities.class.getName()).log(Level.SEVERE, null, ex);
        } finally {
            try {
              
                conn.rs.close();
                conn.st.close();
              out.close();
            } catch (SQLException ex) {
                Logger.getLogger(loadFacilities.class.getName()).log(Level.SEVERE, null, ex);
            }
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
