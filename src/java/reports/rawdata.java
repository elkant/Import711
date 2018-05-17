/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package reports;

import reports.IdGenerator;
import db.dbConn;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Geofrey Nyabuto
 */
public class rawdata extends HttpServlet {

    protected void processRequest(HttpServletRequest request, HttpServletResponse response)
            throws ServletException, IOException {
        String monthrange = "";
        String periodlabel = "";
        String period = "Month"; //The default one
        String periodicgroupby=" "; //note that in the current query there is  an existing group by. Therefore this will be an extra infor to be added on the existing group by
        String isgroupby="yes";
        try {
            response.setContentType("text/html;charset=UTF-8");

//a page to get Report of all the servlets
            String months[] = null;

            String year = "2017";
            String month = "";
            String county = "";
            String form = "moh731";

            if (request.getParameter("year") != null) {
                year = request.getParameter("year");
            }

            if (request.getParameter("county") != null) {
                county = request.getParameter("county");
            }

            if (request.getParameter("month") != null) {
                months = request.getParameterValues("month");
            }

            if (request.getParameter("form") != null) {
                form = request.getParameter("form");
            }
            if (request.getParameter("groupby") != null) {
                isgroupby = request.getParameter("groupby");
            }
            
            
            if(form.equalsIgnoreCase("VMMC")){
                form="vmmc_new";
            }
            
            String pivotform = form;
            if (form.equalsIgnoreCase("MOH 731")) {
                form = "MOH731";
            }
            if (form.equalsIgnoreCase("MOH 711A")) {
                form = "MOH711";
            }
            if (form.equalsIgnoreCase("moh711new")) {
                form = "moh711_new";
            }
            String facilitywhere = "";
            String yearwhere = "";
            String monthwhere = "";
            String countywhere = "";
            String districtwhere = "";
            String reporttype = "";
            String subcounty_countywhere = "";
            String indicatorslist="all";
            
            String sections="all";
            String subsections="all";
            
            String indicatorswhere="";
            
            System.out.println("form : "+form);
//________________________________________________________________________________________________________________________________________________________            
//________________________________________________________________________________________________________________________________________________________
            String duration = "";
            String semi_annual = "";
            String quarter = "";

//==================================================================================================
//XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
            String facil = "361";

            String yearmonthstart = "";
            String yearmonthend = "";
//=====================================================================================================

            String header = "";

            String reportType = "";

            dbConn conn = new dbConn();

            if (request.getParameter("reportType") != null) {
                reportType = request.getParameter("reportType");
            }
            if (request.getParameter("indicators") != null) {
                indicatorslist = request.getParameter("indicators");
            }
            
                //--------------------------sections------------
            String sectionid[]=null;
            
            
            if (request.getParameterValues("sections") != null) 
            {
                sectionid = request.getParameterValues("sections");
            }
            
            String sectionvals="(";
            
            if (request.getParameterValues("sections") != null) 
            {
            for (int a=0;a<sectionid.length;a++){
                if(a==sectionid.length-1){
                     sectionvals+=sectionid[a]+"";
                                         }
                else {
                sectionvals+=sectionid[a]+",";
                }
                                                }
            }
                        sectionvals+=")";
            
            
           
            
            if(sectionvals.equals("()"))
            {
            indicatorswhere+=" ";
            }
            else  
            {
            
            indicatorswhere+=" and sectionid in "+sectionvals+" ";
            
            }
            
            
            //______________________________________________________subsections_______________________________________
            
             String subsectionid[]=null;
            
            
            if (request.getParameterValues("subsection") != null) 
            {
                subsectionid = request.getParameterValues("subsection");
            }
            
            String subsectionvals="(";
             if (request.getParameterValues("subsection") != null) 
            {
            for (int a=0;a<subsectionid.length;a++){
                if(a==subsectionid.length-1){
                     subsectionvals+=subsectionid[a]+"";
                                         }
                else {
                subsectionvals+=subsectionid[a]+",";
                }
                                                }
            }
                        subsectionvals+=")";
            
            
           
            
            if(subsectionvals.equals("()"))
            {
            indicatorswhere+=" ";
            }
            else  
            {
            
            indicatorswhere+=" and subsectionid in "+subsectionvals+" ";
            
            }
            
            //______________________________________________________subsections_______________________________________
            
            
            
            //add sections
            //special indicators
            
             if(indicatorslist.equals("special")) 
             {
            
            indicatorswhere+=" and specialindicator='1'";
            
            }
            
            
            
            
            String reportDuration = "";

            if (request.getParameter("reportDuration") != null) {
                reportDuration = request.getParameter("reportDuration");
            }

            if (request.getParameter("facility") != null && reportType.equals("2")) {
                try {
                    facil = request.getParameter("facility");

                    String getfacil = "select SubPartnerNom,CentreSanteId as mflcode from subpartnera where SubPartnerID='" + facil + "'";
                    conn.rs = conn.st.executeQuery(getfacil);

                    while (conn.rs.next()) {

                        header += " FACILITY : " + conn.rs.getString(1).toUpperCase() + "     MFL CODE  :  " + conn.rs.getString(2) + "  ";

                    }
                } catch (SQLException ex) {
                    Logger.getLogger(rawdata.class.getName()).log(Level.SEVERE, null, ex);
                }

            }

            if (request.getParameter("county") != null && reportType.equals("2")) {
                try {
                    county = request.getParameter("county");
                    subcounty_countywhere = " (county.CountyID='" + county + "') and ";//20160711   
                    String getcounty = "select County from county where CountyID='" + county + "'";
                    conn.rs = conn.st.executeQuery(getcounty);

                    while (conn.rs.next()) {

                        header += " COUNTY : " + conn.rs.getString(1).toUpperCase() + " ";

                    }
                } catch (SQLException ex) {
                    Logger.getLogger(rawdata.class.getName()).log(Level.SEVERE, null, ex);
                }

            }

            int yearcopy = Integer.parseInt(year);
            String yearmonth = "" + year;
            int prevYear = yearcopy - 1;
            int maxYearMonth = 0;
            int monthcopy = 0;
            
            String currentperiodlabel="";//a variable to hold the name of a period inside a conn.rs if .
           

            header += " YEAR : " + year + "";

//        GET REPORT DURATION============================================
//annually
 //____________________________________________________________________________________________________________Annual____________________________________
            if (reportDuration.equals("1")) {
                
                yearmonth = "Annual Report For " + year;
                duration = " and " + form + ".yearmonth BETWEEN '" + prevYear + "10' AND '" + year + "09'";
                 
                period="Year";
                
                periodicgroupby=", period ";
                 monthrange = year;
                            
// tbstatduration="year='"+year+"'";
                                           }
  //____________________________________________________________________________________________________________Semi_annual_____________________________          
            else if (reportDuration.equals("2")) {
	
                period="Semi-Annual";
               
                periodicgroupby=", period  ";
                
                try {
        String startMonth="", endMonth="";
        String semiannualarray[] = request.getParameterValues("semi_annual");

                    String temporaryheader = " SEMI-ANNUAL";

                    for (int p = 0; p < semiannualarray.length; p++) {

                        semi_annual = semiannualarray[p];

						
						
                        String getperiodname = "SELECT name as semiannual_name,months FROM semi_annual WHERE id='" + semi_annual + "'";
                        conn.rs = conn.st.executeQuery(getperiodname);

                        if (conn.rs.next() == true) {
							
			String monthsinsemiannual[]=conn.rs.getString("months").split(",");
                            currentperiodlabel=conn.rs.getString("semiannual_name");
       
                            //_________________add year at the end of period label						
                            if (semi_annual.equals("1")) {
                                //periodname is in format oct-Mar we will make it to be in format Oct 2015-Mar 2016
                                
                                currentperiodlabel = conn.rs.getString("semiannual_name").replace("-", " " + prevYear + "-") + " " + year;
	
							}
							else{
						//periodname is in format Apr-Sep we will make it to be in format Apr-Sep 2016		
						currentperiodlabel=conn.rs.getString("semiannual_name")+" "+year;	
								
							}
						//____________________	
							
                            if(p==0)
                            {
							
							startMonth = monthsinsemiannual[0];
								
                                                        endMonth = monthsinsemiannual[5];
							
							 monthrange =currentperiodlabel;	
								
							
							
                            }
                            if(p==semiannualarray.length-1 && semiannualarray.length>1 )
                            {
								//last row
                        
							 monthrange +=" to "+currentperiodlabel;	
								
								
							
							
							//by now we expect monthrange to be something like Oct-Mar-Apr-Sep
							
							endMonth = monthsinsemiannual[5];
							
                            }
                            
                            if (periodlabel.equals("")) 
                            {
                                periodlabel =conn.rs.getString("semiannual_name");
                            } 
                            
			    else {
								
                                periodlabel += "_" + conn.rs.getString("semiannual_name");
								
                                 }
                        }//end of conn.
                        //get the yearmonthstart date

                        if (p == 0 && semiannualarray.length > 1) {
                             //if the quarters selected are several
                            //this is the starting quarter											 
                            temporaryheader += "S " + currentperiodlabel + " To ";

                            if (new Integer(semi_annual) ==1) {
                                //the year will be deducted by one if the selected months are 10, 11, 12
                                yearmonthstart = " " + form + ".yearmonth  between '" + prevYear + ""+startMonth+"' and " ;

                            } else {
								
								
                                yearmonthstart = " " + form + ".yearmonth  between '" + year + ""+startMonth+"' and ";

                            }

                        } 
                        
                        else if (p == semiannualarray.length - 1 && semiannualarray.length > 1) {
                            //the last month 
                            temporaryheader += currentperiodlabel; //by now we expct something like from Semi-annuals oct-Mar to Apr -Sep
                           //this assumes that the last month can never be 
                            yearmonthend = "'"+year + "" + endMonth+"'";

                        } 
                        
	               else if (p == 0 && semiannualarray.length == 1) {
                            // the number of quarters selected is one	
                            temporaryheader += " " + currentperiodlabel + " ";

                            if (new Integer(semi_annual) == 1) {

                                yearmonthstart = " " + form + ".yearmonth between '" + prevYear + startMonth+"' and";

                            } else {

                                yearmonthstart = " " + form + ".yearmonth between '" + year + "" + startMonth+"' and ";

                            }
                            yearmonthend = "'" + year + "" + endMonth+"'"; //this assumes no end month date can come in oct-dec ..all the year month ends come on mar or sep

                        }//end of last monthly row

                        
                        
                        
                        
                    }//end of for loop

                    header += " " + temporaryheader + "";

                } catch (SQLException ex) {
                    System.out.println(ex);
                }

                if (!semi_annual.equals("")) {

                    duration = " and " + yearmonthstart + "" + yearmonthend;

                }

            } 
            
            // ____________________________________________________________________________________________________Quarterly____________________
            
            else if (reportDuration.equals("3")) {
                period="Quarter";
               
                periodicgroupby=", period ";
                
                try {
        String startMonth="", endMonth="";
        String quarterarray[] = request.getParameterValues("quarter");

                    String temporaryheader = " QUARTER";

                    for (int p = 0; p < quarterarray.length; p++) {

                        quarter = quarterarray[p];

						
						
                        String getMonth = "SELECT months,name as qtrname FROM quarter WHERE id='" + quarter + "'";
                        conn.rs = conn.st.executeQuery(getMonth);

                        if (conn.rs.next() == true) {
							
							String monthsinqtr[]=conn.rs.getString(1).split(",");
                            
                              //_________________add year at the end of period label						
							if (quarter.equals("1")) {
						//periodname is in format oct-Mar we will make it to be in format Oct 2015-Mar 2016
						currentperiodlabel =conn.rs.getString("qtrname").replace("-"," "+prevYear+"-")+" "+year;	
								
							}
							else{
						//periodname is in format Apr-Sep we will make it to be in format Apr-Sep 2016		
						currentperiodlabel=conn.rs.getString("qtrname")+" "+year;	
								
							}
			      //____________________
							
                            if(p==0)
                            {
							
			    startMonth = monthsinqtr[0];
								
                            monthrange = conn.rs.getString("qtrname");
			
                            endMonth = monthsinqtr[2];
                            }
                            if(p==quarterarray.length-1 && quarterarray.length>1 )
                            {
								//last row
                            monthrange +=" to "+conn.rs.getString("qtrname");
							
			    endMonth = monthsinqtr[2];
							
                            }
                            
                            if (periodlabel.equals(""))// note period label gets all the periods in my loop 
                            {
                                periodlabel = conn.rs.getString("qtrname");
                            } 
                            
			   else {
								
                                periodlabel += "_" + conn.rs.getString("qtrname");
								
                                 }
                        }//end of if
                        //get the yearmonthstart date

                        if (p == 0 && quarterarray.length > 1) {
                             //if the quarters selected are several
                            //this is the starting quarter											 
                            temporaryheader += "S " + currentperiodlabel + " To ";

                            if (new Integer(quarter) ==1) {
                                //the year will be deducted by one if the selected months are 10, 11, 12
                                yearmonthstart = " " + form + ".yearmonth  between '" + prevYear + ""+startMonth+"' and" ;

                            } else {
								
								
                                yearmonthstart = " " + form + ".yearmonth  between '" + year + ""+startMonth+"' and";

                            }

                        }
                        else if (p == quarterarray.length - 1 && quarterarray.length > 1) {
                            //the last month 
                            temporaryheader += currentperiodlabel; //by now we expct something like from Quarters oct-Dec to Jan -Mar

                            yearmonthend = "'"+year + "" + endMonth+"'";

                        } 
			else if (p == 0 && quarterarray.length == 1) {
                            // the number of quarters selected is one	
                            temporaryheader += " " + currentperiodlabel + " ";

                            if (new Integer(quarter) == 1) {

                                yearmonthstart = " " + form + ".yearmonth between '" + prevYear + startMonth+"' and ";
                                
                                yearmonthend = " '" + prevYear + "" + endMonth+"' ";
                            } else {

                                yearmonthstart = " " + form + ".yearmonth between '" + year + "" + startMonth+"' and ";
                                
                                yearmonthend = " '" + year + "" + endMonth+"' ";
                            }
                           

                        }//end of last monthly row

                        
                        
                        
                        
                    }//end of for loop

                    header += " " + temporaryheader + "";

                } catch (SQLException ex) {
                    System.out.println(ex);
                }

                if (!quarter.equals("")) {

                    duration = " and " + yearmonthstart + "" + yearmonthend;

                }

            }
            
     //_______________________________________________________________________________________________________________monthly______________________________       
            else  if (reportDuration.equals("4")) {
                
                period="Month";
               
                periodicgroupby=", period ";
                
                try {

                    months = request.getParameterValues("month");

                    String temporaryheader = " MONTH";

                    for (int u = 0; u < months.length; u++) {

                        month = months[u];

                        String getMonth = "SELECT name FROM month WHERE id='" + month + "'";
                        conn.rs = conn.st.executeQuery(getMonth);

                        if (conn.rs.next() == true) {
                            
                          
                        //_________________add year at the end of period label						
							if (new Integer(month)>=10 && new Integer(month) <= 12) {
						//periodname is in format oct-Mar we will make it to be in format Oct 2015-Mar 2016
						currentperiodlabel =conn.rs.getString("name").substring(0, 3)+" "+prevYear;	
								
							}
							else{
						//periodname is in format Apr-Sep we will make it to be in format Apr-Sep 2016		
						currentperiodlabel=conn.rs.getString("name").substring(0, 3)+" "+year;	
								
							}
			      //____________________    
                            
                            
                            if(u==0)
                            {
                            monthrange = conn.rs.getString(1);
                            }
                            if(u==months.length-1 && months.length>1 )
                            {
                            monthrange +=" to "+conn.rs.getString(1);
                            }
                            
                            if (periodlabel.equals("")) 
                            {
                                periodlabel = conn.rs.getString("name");
                            } 
                            else {
                                periodlabel += "_" + conn.rs.getString("name");
                            }
                                                      }
                        //get the yearmonthstart date

                        if (u == 0 && months.length > 1) {
                            //if the month selected are several
                            //this is the starting month											 
                            temporaryheader += "S " + currentperiodlabel + " To ";

                            if (new Integer(month) >= 10) {
                                
                                //the year will be deducted by one if the selected months are 10, 11, 12
                                yearmonthstart = " " + form + ".yearmonth  between '" + prevYear + "" + month+"' and ";

                            } else {
                                yearmonthstart = " " + form + ".yearmonth  between '" + year + "0" + month+"' and ";

                            }

                        } else if (u == months.length - 1 && months.length > 1) {
                            //the last month 
                            

                            
                            if (new Integer(month) >= 10) {

                                yearmonthend ="'"+prevYear + month+"'";

                            } else {

                                yearmonthend = "'"+year + "0" + month+"'";

                            }

                        } else if (u == 0 && months.length == 1) {
                            // the number of months selected is one	
                            temporaryheader += " " + currentperiodlabel + " ";

                            if (new Integer(month) >= 10) {

                                yearmonthstart = " " + form + ".yearmonth = " + prevYear + month;

                            } else {

                                yearmonthstart = " " + form + ".yearmonth = " + year + "0" + month;

                            }
                            yearmonthend = "";

                        }//end of last monthly row

                        
                        
                        
                        
                    }//end of for loop

                    header += " " + temporaryheader + "";

                } catch (SQLException ex) {
                    Logger.getLogger(rawdata.class.getName()).log(Level.SEVERE, null, ex);
                }

                if (!month.equals("")) {

                    duration = " and " + yearmonthstart + "" + yearmonthend;

                }

            }//end of monthly records 
            
            else {
                duration = "";
                 }

            //if someone doest want periodic grouping, then dont add a group by in the report.. get numbers for all rows. 
            
            if(isgroupby.equals("No")){
            
            periodicgroupby="";
            
            }
            
//______________________________________________________________________________________COUNTY , SUBCOUNTY AND 

            String subcountywhere = "";

            String subcounty = "";

            if( reportType.equalsIgnoreCase("2"))
            { //get data per county if the selected
                
                
            if (!request.getParameter("subcounty").equals("")){

                subcounty = request.getParameter("subcounty");
                subcounty_countywhere = " (district.DistrictID='" + subcounty + "') and ";
            }

            if (!request.getParameter("county").equals("")) {
                county = request.getParameter("county");
                subcounty_countywhere = " (district.countyid='" + request.getParameter("county") + "') and ";//20160711 
            }

           

            if (!county.equals("")) {

                countywhere = " and district.countyid = '" + county + "'";

            }

            if (!subcounty.equals("")) 
            {

                subcountywhere = " and subpartnera.DistrictID = '" + subcounty + "'";

            }

            if (!facil.equals("") && reportType.equalsIgnoreCase("2")) 
            {

                facilitywhere = " and " + form + ".facility = '" + facil + "'";

            }
            
        }

            String joinedwhwere = " where subpartnerid not like '%summary%' " + duration + " " + countywhere + " " + subcountywhere+" "+facilitywhere;

           // System.out.println(""+joinedwhwere);
            //we need a case statement in our main query. This will allow for friendly display of 
            
        String myperiodcase="";
        
        if(isgroupby.equals("No")){
            
            myperiodcase="'"+monthrange+"' as period";
            
            }
        
        else {
        
        if(period.equalsIgnoreCase("Year"))
        {
        
        myperiodcase=" case when Annee !='' then Annee else 'no year' end as period ";
        
        }
        else if(period.equalsIgnoreCase("Semi-Annual")){
        
        myperiodcase=" case when (SUBSTRING(Yearmonth,5,6) >=01 and SUBSTRING(Yearmonth,5,6) <=03) || (SUBSTRING(Yearmonth,5,6) >=10 and SUBSTRING(Yearmonth,5,6) <=12) then 'Oct-Mar' " +
"when SUBSTRING(Yearmonth,5,6) >=04 and SUBSTRING(Yearmonth,5,6) <=09 then 'Apr-Sep' else 'No period' end as period ";
        }
        else if(period.equalsIgnoreCase("Quarter")){
        
        myperiodcase=" case when (SUBSTRING(Yearmonth,5,6) >=01 and SUBSTRING(Yearmonth,5,6) <=03) then 'Jan-Mar'" +
"when  (SUBSTRING(Yearmonth,5,6) >=10 and SUBSTRING(Yearmonth,5,6) <=12) then 'Oct-Dec' " +
"when SUBSTRING(Yearmonth,5,6) >=04 and SUBSTRING(Yearmonth,5,6) <=06 then 'Apr-Jun' " +
"when SUBSTRING(Yearmonth,5,6) >=07 and SUBSTRING(Yearmonth,5,6) <=09 then 'Jul-Sep' " +
"else 'No period' end as period  ";
        }
        else if (period.equalsIgnoreCase("Month")){
        
        myperiodcase=" case " +
" when SUBSTRING(Yearmonth,5,6)= '01' then 'Jan' " +
" when SUBSTRING(Yearmonth,5,6)= '02' then 'Feb' " +
" when SUBSTRING(Yearmonth,5,6)= '03' then 'Mar' " +
" when SUBSTRING(Yearmonth,5,6)= '04' then 'Apr' " +
" when SUBSTRING(Yearmonth,5,6)= '05' then 'May' " +
" when SUBSTRING(Yearmonth,5,6)= '06' then 'Jun' " +
" when SUBSTRING(Yearmonth,5,6)= '07' then 'Jul' " +
" when SUBSTRING(Yearmonth,5,6)= '08' then 'Aug' " +
" when SUBSTRING(Yearmonth,5,6)= '09' then 'Sep' " +
" when SUBSTRING(Yearmonth,5,6)= '10' then 'Oct' " +
" when SUBSTRING(Yearmonth,5,6)= '11' then 'Nov' " +
" when SUBSTRING(Yearmonth,5,6)= '12' then 'Dec' "
+ " else 'No period' end as period ";
        
        }
        
        }
        
            
//_________________________________________________________________END OF FILTERING SECTION_______________________________________________________________            
//________________________________________________________________________________________________________________________________________________________            
//an array to store haeder information.
//the header information should appear only if certain parameters are met
//The parameters listed in here can be removed if the report type doesnt require some parameters
            ArrayList Headerorgunits = new ArrayList();
            Headerorgunits.add(period);
            Headerorgunits.add("County");
            Headerorgunits.add("Sub-County");
            Headerorgunits.add("Facility");
            Headerorgunits.add("MFL Code");
//An arralist to store a list of columns that will be selected from the database
            ArrayList dbcolumns = new ArrayList();
            ArrayList shortlabels = new ArrayList();

            ArrayList labels = new ArrayList();

            ArrayList tablename = new ArrayList();

            ArrayList iscumulative = new ArrayList();

            ArrayList ispercent = new ArrayList();

            // ArrayList isactive=new ArrayList();
            //An arralist to store a list of worksheets that will be selected from the sections and the respective service area to determine the facilities whose data will appear in that sheet
            ArrayList worksheets = new ArrayList();
          //An arralist to store distinct worksheets. This will be derived from the the sections column
            ArrayList distinctsheets = new ArrayList();
            ArrayList distinctservicearea = new ArrayList();

//create an array to store the number of row for each excel worksheet. 
//This will help in retrieving the number of rows for each month since we are wring data for different months with increasing rows.
//the size of that array will be determined by the number of excel worksheets
            String qrform = form;
            if(!form.equalsIgnoreCase("vmmc_new")){
             qrform=form.replace("_", "");   
            }
            
            String selectdistinctworksheet = "select section,servicearea from pivottable where form='" + qrform + "' and active='1' "+indicatorswhere+" group by section order by order_per_form";

            conn.rs = conn.st.executeQuery(selectdistinctworksheet);
            System.out.println("worksheetquery: "+selectdistinctworksheet);
            while (conn.rs.next()) {
//add the name of distinct sections
                distinctsheets.add(conn.rs.getString(1).replace("/", "_"));

                String servicearea = "  2=2 ";
                if (conn.rs.getString(2) != null) {
                    //servicearea = "  " + conn.rs.getString(2) + "=1";
                }
                distinctservicearea.add(servicearea);

                                   }

            int rowstartpersheet[] = new int[distinctsheets.size()];

//initialize the row start position for each workshett with 2
            for (int x = 0; x < rowstartpersheet.length; x++) 
            {

                rowstartpersheet[x] = 2;

                
            }
            
            qrform = form;
            if(!form.equalsIgnoreCase("vmmc_new")){
             qrform=form.replace("_", "");   
            }
            
            String getattribs = "select indicator,label,section,cumulative,percentage,active ,shortlabel from pivottable where form='" + qrform + "' "+indicatorswhere+" order by order_per_form, section";
            conn.rs = conn.st.executeQuery(getattribs);
            System.out.println("get attributes : "+getattribs);
            while (conn.rs.next()) {

//add active indicators only
                if (conn.rs.getString("active").equals("1")) {
//System.out.println(conn.rs.getString("indicator")+"");
//add indicator
                    dbcolumns.add(conn.rs.getString("indicator"));
                    shortlabels.add(conn.rs.getString("shortlabel"));
//add label
                    if (form.equalsIgnoreCase("MOH731")) {
                        
                        labels.add(conn.rs.getString("shortlabel") + " \n" + conn.rs.getString("label"));

                    } else {
                        labels.add(conn.rs.getString("shortlabel"));
                    }
//add worksheets
                    worksheets.add(conn.rs.getString("section").replace("/", "_"));

                    String perc = "0";
                    String cum = "0";

                    if (conn.rs.getString("cumulative") != null) {
                        iscumulative.add(conn.rs.getString("cumulative"));
                    } else {
                        iscumulative.add(cum);
                    }

                    if (conn.rs.getString("percentage") != null) {
                        ispercent.add(conn.rs.getString("percentage"));
                    } else {
                        ispercent.add(perc);
                    }

                }//end of active 

            }  //end of pivot table active
//labels.add("ART High Volume");
//labels.add("HTC High Volume");
//labels.add("PMTCT High Volume");

//System.out.println(perfacilselect);
//______________________________________________________________________________________
//                       CREATE THE WORKSHEETS          
//______________________________________________________________________________________  
            XSSFWorkbook wb = new XSSFWorkbook();

            XSSFFont font = wb.createFont();
            font.setFontHeightInPoints((short) 18);
            font.setFontName("Cambria");
            font.setColor((short) 0000);
            CellStyle style = wb.createCellStyle();
            style.setFont(font);
            style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
            XSSFFont font2 = wb.createFont();
            font2.setFontName("Cambria");
            font2.setColor((short) 0000);
            CellStyle style2 = wb.createCellStyle();
            style2.setFont(font2);
            style2.setBorderTop(HSSFCellStyle.BORDER_THIN);
            style2.setBorderBottom(HSSFCellStyle.BORDER_THIN);
            style2.setBorderLeft(HSSFCellStyle.BORDER_THIN);
            style2.setBorderRight(HSSFCellStyle.BORDER_THIN);
            style2.setAlignment(HSSFCellStyle.ALIGN_LEFT);

            XSSFCellStyle stborder = wb.createCellStyle();
            stborder.setBorderTop(HSSFCellStyle.BORDER_THIN);
            stborder.setBorderBottom(HSSFCellStyle.BORDER_THIN);
            stborder.setBorderLeft(HSSFCellStyle.BORDER_THIN);
            stborder.setBorderRight(HSSFCellStyle.BORDER_THIN);
            stborder.setAlignment(HSSFCellStyle.ALIGN_CENTER);

            XSSFCellStyle stylex = wb.createCellStyle();
            stylex.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);
            stylex.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            stylex.setBorderTop(HSSFCellStyle.BORDER_THIN);
            stylex.setBorderBottom(HSSFCellStyle.BORDER_THIN);
            stylex.setBorderLeft(HSSFCellStyle.BORDER_THIN);
            stylex.setBorderRight(HSSFCellStyle.BORDER_THIN);
            stylex.setAlignment(HSSFCellStyle.ALIGN_LEFT);

            XSSFCellStyle stylesum = wb.createCellStyle();
            stylesum.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);
            stylesum.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            stylesum.setBorderTop(HSSFCellStyle.BORDER_THIN);
            stylesum.setBorderBottom(HSSFCellStyle.BORDER_THIN);
            stylesum.setBorderLeft(HSSFCellStyle.BORDER_THIN);
            stylesum.setBorderRight(HSSFCellStyle.BORDER_THIN);
            stylesum.setAlignment(HSSFCellStyle.ALIGN_CENTER);

            XSSFFont fontx = wb.createFont();
            fontx.setColor(HSSFColor.BLACK.index);
            fontx.setFontName("Cambria");
            stylex.setFont(fontx);
            stylex.setWrapText(true);

            stylesum.setFont(fontx);
            stylesum.setWrapText(true);
            int cellrange[] = new int[4];

            for (int b = 0; b < distinctsheets.size(); b++) {
                
                XSSFSheet shet = wb.createSheet(distinctsheets.get(b).toString().toUpperCase());

              //Made my life veery simple...
                shet.setDisplayGridlines(false);
                shet.createFreezePane(5, 2);
                int firstcell = 1;


//create headers for that worksheet
                XSSFRow rw = shet.createRow(1);
                int headercellpos = 0;
//create the orgunit header eg MONTH | COUNTY | SUBCOUNTY  | FACILITY

                for (int e = 0; e < Headerorgunits.size(); e++) {
                    XSSFCell cell0 = rw.createCell(headercellpos);
                    cell0.setCellValue(Headerorgunits.get(e).toString());
                    cell0.setCellStyle(stylex);

                    headercellpos++;
//shet.setColumnWidth(e, 6000);  

                }

//create the indicators header eg HV0101 | HIV 09676  | TOTAL    
                for (int c = 0; c < dbcolumns.size(); c++) {
//compare if the indicator belongs to the specified section and hence worksheet 
//recall, each indicator has got an associated section / worksheet
//An indicator should be put as an header in the respective worksheet
                    if (worksheets.get(c).equals(distinctsheets.get(b))) {

                        shet.setColumnWidth(headercellpos, 6000);
                        XSSFCell cell0 = rw.createCell(headercellpos);
                        cell0.setCellValue(labels.get(c).toString());
                        cell0.setCellStyle(stylex);
                        headercellpos++;
                    }//end of comparing if

                }//end of for loop

//create is validated header
//highvolume
                shet.setColumnWidth(headercellpos, 6000);
                XSSFCell cell0x = rw.createCell(headercellpos);
                cell0x.setCellValue("FPRH");
                cell0x.setCellStyle(stylex);
                headercellpos++;
//FPRH			

                shet.setColumnWidth(headercellpos, 6000);
                XSSFCell cell0 = rw.createCell(headercellpos);
                cell0.setCellValue("MCH");
                cell0.setCellStyle(stylex);
                headercellpos++;

                shet.setColumnWidth(headercellpos, 6000);
                XSSFCell cell1 = rw.createCell(headercellpos);
                cell1.setCellValue("Training");
                cell1.setCellStyle(stylex);
                headercellpos++;

                shet.setColumnWidth(headercellpos, 6000);
                XSSFCell cell2 = rw.createCell(headercellpos);
                cell2.setCellValue("WASH");
                cell2.setCellStyle(stylex);
                headercellpos++;
                
//                shet.setColumnWidth(headercellpos, 6000);
//                XSSFCell cell2a = rw.createCell(headercellpos);
//                cell2a.setCellValue("GSN");
//                cell2a.setCellStyle(stylex);
//                headercellpos++;
                
                shet.setColumnWidth(headercellpos, 6000);
                XSSFCell cell2b = rw.createCell(headercellpos);
                cell2b.setCellValue("Latitude");
                cell2b.setCellStyle(stylex);
                headercellpos++;
                
                shet.setColumnWidth(headercellpos, 6000);
                XSSFCell cell2c = rw.createCell(headercellpos);
                cell2c.setCellValue("Longitude");
                cell2c.setCellStyle(stylex);
                headercellpos++;
                
                shet.setColumnWidth(headercellpos, 6000);
                XSSFCell cell2d = rw.createCell(headercellpos);
                cell2d.setCellValue("Ward");
                cell2d.setCellStyle(stylex);
                headercellpos++;

                shet.setColumnWidth(headercellpos, 6000);
                XSSFCell cell3 = rw.createCell(headercellpos);
                cell3.setCellValue("Is subcounty Summary ?");
                cell3.setCellStyle(stylex);
                headercellpos++;

            }

//______________________________________________________________________________________
//______________________________________________________________________________________
//--------------------------------------------------------------------------------------------
//             MONTH LOOPS
//-------------------------------------------------------------------------------------------- 
            int cumulativestatingpoint = 3;
            boolean cumstartpointnoted = false;
            int colposcopy = 0;
            //for (int w = 0; w < months.length; w++) {

                

                String perfacilselect = "select " + myperiodcase + " , CONCAT(UPPER(SUBSTRING(County,1,1)),LOWER(SUBSTRING(County,2))) as County , DistrictNom as District , subpartnera.SubPartnerNom as facility ,CentreSanteId as mflcode ,  district.CountyID as countyid , ";
   
                //have a stringto get all distinct periods before execution of the main query. This will help in knowing the last
                String getdistinctperiod="select " + myperiodcase;
//--------------------------------------------------------------------------------------------
//             PREPARE SELECT
//--------------------------------------------------------------------------------------------
//prepare selects
System.out.println("element size : "+dbcolumns.size());
                for (int a = 0; a < dbcolumns.size(); a++) {

//if the indicator is a percent, get an avaerage
                    if (ispercent.get(a).equals("1")) {
                        perfacilselect += "  AVG(" + dbcolumns.get(a) + ") as '" + shortlabels.get(a)+"'";

                    } else if (iscumulative.get(a).equals("1")) {
                        perfacilselect += "  SUBSTRING_INDEX(GROUP_CONCAT(CAST(IFNULL(" + dbcolumns.get(a) + ",0) AS CHAR) ORDER BY yearmonth DESC),',',1) as '" + shortlabels.get(a)+"'";
//SUBSTRING_INDEX(GROUP_CONCAT(CAST(IFNULL(HV0303,0) AS CHAR) ORDER BY yearmonth DESC),',',1)
                        
                    } else {
                        perfacilselect += "  SUM(IFNULL(" + dbcolumns.get(a) + ",0)) as '" + shortlabels.get(a)+"'";

                    }

//if the item is not the last, append a comma
//if(a<dbcolumns.size()-1){
                    perfacilselect += " ,";

// } 
System.out.println("enntered loop");
                }
//			

//---------------------------------add highvolume,gsn,latitude,ward------------------------------------------------
                perfacilselect += " IFNULL(FPRH,0) AS 'FPRH', IFNULL(MCH,0) as MCH,  IFNULL(Training,0) as Training,  IFNULL(WASH,0) as WASH,IFNULL(latitude,0) as latitude,IFNULL(longitude,0) as longitude,IFNULL(ward,0) as ward,";

//-------------------------------------------------------------------------------------------------
//     FROM  
//------------------------------------------------------------------------------------  
                perfacilselect += " subpartnera.SubPartnerID as SubPartnerID  from " + form + "  left join ( subpartnera left join (district join county on county.CountyID=district.CountyID ) on district.DistrictID = subpartnera.DistrictID )  on " + form + ".facility = subpartnera.SubPartnerID ";
 
                getdistinctperiod+= " from " + form + "  left join ( subpartnera left join (district left join county on county.CountyID=district.CountyID ) on district.DistrictID = subpartnera.DistrictID )  on " + form + ".facility = subpartnera.SubPartnerID ";
                
//------------------------------------------------------------------------------------------
// WHERE 
//------------------------------------------------------------------------------------------ 
                perfacilselect += joinedwhwere; //contains any filterings
               
                getdistinctperiod += joinedwhwere; //contains any filterings

//-----------------------------------------------------------------------------------------
//GROUP BY 
//----------------------------------------------------------------------------------------
                
                perfacilselect += " group by subpartnera.SubPartnerNom "+periodicgroupby+" order by yearmonth,county ";
                
                getdistinctperiod += " group by period order by yearmonth,county ";

                
                
                String lastperiod="";
                ArrayList alldistinctperiods = new ArrayList();
            //System.out.println(""+getdistinctperiod);                
                
                conn.rs = conn.st.executeQuery(getdistinctperiod);
                System.out.println("perfacilselect: "+perfacilselect);
                System.out.println("getdistinctperiod: "+getdistinctperiod);
                while (conn.rs.next() == true) 
                {                    
                 lastperiod = conn.rs.getString(1);    //here am asumming the last period will appear last and so will be mantained in the                
                alldistinctperiods.add(lastperiod);
                }
                
                
                String sectioncopy = "";

                int sheetpos = 0;
                int rowpos = 2;

//-----------------INSIDE THE DATA FORM---------------------------------
//if the section (eg HTC, PMTCT) changes, change the current workshhet index too
//also, reset the row position counter to begin from 2 again. 
                XSSFSheet shet = null;

//      if(--!sectioncopy.equals(shet)){}
//create the org unit data values e.g BARINGO | BARINGO CENTRAL |KABARNET DISTRICT HOSPITAL | MFL CODE
                for (int g = 0; g < distinctsheets.size(); g++) {

                   

                    rowpos = rowstartpersheet[g];

                    shet = wb.getSheetAt(g);
                    int colpos = 0;

                    String finalquery = perfacilselect.replace("1=1", distinctservicearea.get(g).toString());
                    System.out.println("final query" + finalquery);
                    conn.rs = conn.st.executeQuery(finalquery);
                    while (conn.rs.next()) {

if (lastperiod.equalsIgnoreCase(conn.rs.getString("period")) && cumstartpointnoted == false) {
//save the current row position.
//get it from the current rowstartpersheet[g]
                        
                        cumulativestatingpoint = rowstartpersheet[0];

                        cumstartpointnoted = true;                         
//this ensures that we are fetching dataonce.
// System.out.println("LAST STARTING POINT__"+cumulativestatingpoint);
                        

                    }
                        
                        //the fourth cell should     
                        XSSFRow rw = shet.createRow(rowpos);
                        for (int e = 0; e < Headerorgunits.size(); e++) {
                            XSSFCell cell0 = rw.createCell(colpos);
                            //for mfl code, last header, output integers
                            if (e > 3) {
                                cell0.setCellValue(conn.rs.getInt(e + 1));
                            } else {
                                cell0.setCellValue(conn.rs.getString(e + 1));
                            }

                            cell0.setCellStyle(style2);
                            colpos++;

                        }

                        //_________________________________________________________________
                        //VALUES
                        //_________________________________________________________________
                        //create the indicators values eg 90 | 45  | 356    
                        for (int c = 0; c < dbcolumns.size(); c++) {
//get the section of the current dbcolumn

//compare if the indicator belongs to the specified section and hence worksheet 
                            //recall, each indicator has got an associated section / worksheet
                            //An indicator should be put as an header in the respective worksheet
                            if (worksheets.get(c).equals(distinctsheets.get(g))) {

                                XSSFCell cell0 = rw.createCell(colpos);
                                cell0.setCellValue(conn.rs.getInt(shortlabels.get(c).toString()));
                                cell0.setCellStyle(stborder);
                                colpos++;
                            }//end of comparing if

                        }//end of for loop

                        //			

                        XSSFCell cell0x = rw.createCell(colpos);
                        cell0x.setCellValue(conn.rs.getInt("FPRH"));
                        cell0x.setCellStyle(stborder);
                        colpos++;

                        //MCH
                        XSSFCell cell0 = rw.createCell(colpos);
                        cell0.setCellValue(conn.rs.getInt("MCH"));
                        cell0.setCellStyle(stborder);
                        colpos++;

                        //Training
                        XSSFCell cell1 = rw.createCell(colpos);
                        cell1.setCellValue(conn.rs.getInt("Training"));
                        cell1.setCellStyle(stborder);
                        colpos++;

                        //WASH
                        XSSFCell cell2 = rw.createCell(colpos);
                        cell2.setCellValue(conn.rs.getInt("WASH"));
                        cell2.setCellStyle(stborder);
                        colpos++;
                        
//                        //PMTCT_highvolume
//                        XSSFCell cell2a = rw.createCell(colpos);
//                        cell2a.setCellValue(conn.rs.getInt("GSN"));
//                        cell2a.setCellStyle(stborder);
//                        colpos++;
//                        
                         //Latitude
                        XSSFCell cell2b = rw.createCell(colpos);
                        cell2b.setCellValue(conn.rs.getString("latitude"));
                        cell2b.setCellStyle(stborder);
                        colpos++;
                        
                         //Longitude
                        XSSFCell cell2c = rw.createCell(colpos);
                        cell2c.setCellValue(conn.rs.getString("longitude"));
                        cell2c.setCellStyle(stborder);
                        colpos++;
                        
                        //Ward
                        XSSFCell cell2d = rw.createCell(colpos);
                        cell2d.setCellValue(conn.rs.getString("ward"));
                        cell2d.setCellStyle(stborder);
                        colpos++;
                        
//county summary
                        String isvalidated = "No";
                        String spid="";
if(conn.rs.getString("facility")!=null){spid=conn.rs.getString("facility");}
                        if (spid.contains("Summary")) {
                            isvalidated = "Yes";
                        }
                        XSSFCell cell3 = rw.createCell(colpos);
                        cell3.setCellValue(isvalidated);
                        cell3.setCellStyle(stborder);

                        colpos++;

                        rowpos++;
                        if (colpos > 1) {
                            colposcopy = colpos - 1;

                        }
                        colpos = 0;
                    }// end of while loop getting data from the db

//____________________________________________________________________________________
//COLUMN AUTOSIZE COLUMN WIDTH , AUTOFILTER & AUTOSUM
//____________________________________________________________________________________
//At this point we are sure this is the last row and we have exhausted fetching data for all periods
//System.out.println(" Column position Before "+colposcopy);
  //                  if (lastperiod.equalsIgnoreCase(conn.rs.getString("period"))) {

//System.out.println(" Column after "+colposcopy);
                        shet.setAutoFilter(new CellRangeAddress(1, rowpos - 1, 0, colposcopy));

    //System.out.println("1,"+rowpos+",0,"+colposcopy);
                        for (int e = 0; e < Headerorgunits.size(); e++) {
                            shet.autoSizeColumn(e);
                        }

                        //autosum
                        
                        XSSFRow initialrow = shet.getRow(2);
                        XSSFRow prevrow = shet.getRow(rowpos - 1);
                        XSSFRow cumrow = shet.getRow(cumulativestatingpoint);//not used for now but let it stay here may be of use in future  
                        
                        XSSFRow rwsum = shet.createRow(rowpos);
                        int colpossum = 0;
                        int firstcols = 5;
                        int periodcolumn = 0;

                        for (int f = 0; f < firstcols; f++) {

                            if (f == 0) {

                                XSSFCell cellsum = rwsum.createCell(0);
                                cellsum.setCellValue("Total");
                                cellsum.setCellStyle(stylesum);
                            } else if (f > 0 && f < firstcols) {
                                XSSFCell cellsum = rwsum.createCell(f);
                                cellsum.setCellValue(" ");
                                cellsum.setCellStyle(stylesum);
                            }
                        }

                        for (int c = 0; c < dbcolumns.size(); c++) {

                            if (worksheets.get(c).equals(distinctsheets.get(g))) {

                                XSSFCell cellsum = rwsum.createCell(colpossum + firstcols);
                                XSSFCell initialcell = initialrow.getCell(colpossum + firstcols);

                                String cellformula = "";
                                XSSFCell prevcell = prevrow.getCell(colpossum + firstcols);//the last cell of the current column. rem whe are looping through all the columns
                                //periodcolumncell
                                XSSFCell initialperiodcell = initialrow.getCell(periodcolumn);// the first cell of the period(month, year) column
                                XSSFCell currentperiodcell = prevrow.getCell(periodcolumn);//the last cell of the period column

                                cellsum.setCellType(cellsum.CELL_TYPE_FORMULA);
                                String startcellreference = initialcell.getReference();
                                String lastavailableperiod = currentperiodcell.getStringCellValue();
                                if (iscumulative.get(c).equals("1")) {
        //initialcell=cumrow.getCell(colpossum+firstcols);
                                    //Create a formula that sums all the data for the last column to appear in the row. This should pick the first column which is the period column.
                                    cellformula = "SUMPRODUCT(--(SUBTOTAL(3,OFFSET(INDEX(" + initialperiodcell.getReference() + ":" + currentperiodcell.getReference() + ",1,1),ROW(" + initialperiodcell.getReference() + ":" + currentperiodcell.getReference() + ")-ROW(INDEX(" + initialperiodcell.getReference() + ":" + currentperiodcell.getReference() + ",1,1)),0))=1),--(" + initialperiodcell.getReference() + ":" + currentperiodcell.getReference() + "=\"" + lastavailableperiod + "\")," + startcellreference + ":" + prevcell.getReference() + ")";

                                } else if (ispercent.get(c).equals("1")) {
        //initialcell=cumrow.getCell(colpossum+firstcols);
                                    //Create a formula that sums all the data for the last column to appear in the row. This should pick the first column which is the period column.
                                    cellformula = "ROUNDUP(SUBTOTAL(9," + startcellreference + ":" + prevcell.getReference() + "),1)"; // round of maximum values

                                } else {

                                    cellformula = "SUBTOTAL(9," + startcellreference + ":" + prevcell.getReference() + ")";
                                }

        //for cumulative indicators, we need to do column total for the last selected month
                                //we therefore need to always track where the previous month started at.
                                cellsum.setCellFormula(cellformula);
                                cellsum.setCellStyle(stylesum);

                                colpossum++;
                            }
                        }
                        //merge last cell
                        shet.addMergedRegion(new CellRangeAddress(rowpos, rowpos, 0, 4));
                   // } end of checking  if this is the last month //disbled for now

//
                    rowstartpersheet[g] = rowpos;

                }// end of distinct sheets report

           // }//end of monthly loop

            IdGenerator IG = new IdGenerator();
            String createdOn = IG.CreatedOn();

            System.out.println("" + form.toUpperCase().trim() + "_RPT_FOR_" + year.trim() + "(" + periodlabel + ")_GEN_" + createdOn.trim() + ".xlsx");

            ByteArrayOutputStream outByteStream = new ByteArrayOutputStream();
            wb.write(outByteStream);
            byte[] outArray = outByteStream.toByteArray();
            response.setContentType("application/ms-excel");
            response.setContentLength(outArray.length);
            response.setHeader("Expires:", "0"); // eliminates browser caching
            response.setHeader("Content-Disposition", "attachment; filename=" + form.toUpperCase().trim() + "_REPORT_FOR_" + year.trim() + "(" + periodlabel.replace(" ","") + ")_CREATED_" + createdOn.trim() + ".xlsx");
            OutputStream outStream = response.getOutputStream();
            outStream.write(outArray);
            outStream.flush();
        } catch (SQLException ex) {
            Logger.getLogger(rawdata.class.getName()).log(Level.SEVERE, null, ex);
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
