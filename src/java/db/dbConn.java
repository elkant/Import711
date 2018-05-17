/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package db;

import java.io.BufferedReader;
import java.io.DataInputStream;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.net.URL;
import java.sql.Connection;
import java.sql.Statement;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.util.logging.Level;
import java.util.logging.Logger;
/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

/**
 *
 * @author SIXTYFOURBIT
 */
public final class dbConn {

    public ResultSet rs0,rs, rs1, rs2, rs3, rs4, rs_1, rs_2, rs_3, rs_4, rs_5, rs_6, anc_sch_rs;
    public Statement st0,st, st1, st2, st3, st4, st_1, st_2, st_3, st_4, st_5, st_6, anc_scheduling_st;
    public Statement state,state1,state2,state3,state4;
    public PreparedStatement pst1,pst2,pst3,pst;
    public Connection connect;
    String mydrive = "";
    public static int issetdbcalled_file_exists = 2;
    public static int issetdbcalled_exception = 2;
    public static int issetdbcalled_wrongpword = 2;
   public  String dbsetup[] = new String[4];

    public dbConn() {
        try {
            Class.forName("com.mysql.jdbc.Driver").newInstance();
            // Connection conn = DriverManager.getConnection("jdbc:mysql://localhost:3306/mhc","root", "");
             connect = null;





            //if the saved host name is less than 2 letters long, then thats not a genuine host name

            URL location = dbConn.class.getProtectionDomain().getCodeSource().getLocation();


            mydrive = location.getFile().substring(1, 2);

            if (getdbsettings(mydrive) == true) {

                //String myfile=getServletContext().getRealPath("/dbsettings.txt");

                if (dbsetup[0] != null) {
//connect = DriverManager.getConnection("jdbc:sqlserver://nkudat01:1433;databaseName=APHIAMAINDB_2014_NEWTOOLS;user=sa;password=Fhiimpact!;");
                    if(dbsetup[3]==null){connect = DriverManager.getConnection("jdbc:mysql://" + dbsetup[0] + "/" + dbsetup[1], dbsetup[2],"");
}
                    else{
                    connect = DriverManager.getConnection("jdbc:mysql://" + dbsetup[0] + "/" + dbsetup[1], dbsetup[2],dbsetup[3]);
                        System.out.println("jdbc:mysql://"+dbsetup[0] + "/" + dbsetup[1]+","+ dbsetup[2]+","+dbsetup[3]);
                    
                    }


                } else {
                    //call the page thats sets up the database
                    //use if to avoid calling the db.jsp twice.
                    if (issetdbcalled_wrongpword %2== 0) {
                        calldbjsp();
                        issetdbcalled_wrongpword ++;
                    }
                    else{
                     issetdbcalled_wrongpword ++;
                    }

                }


                //initialize this three values
                issetdbcalled_exception = 2;
                issetdbcalled_file_exists = 2;
                issetdbcalled_wrongpword = 2;



                st_1 = connect.createStatement();
                st_2 = connect.createStatement();
                st_3 = connect.createStatement();
                st_4 = connect.createStatement();
                st_5 = connect.createStatement();
                st_6 = connect.createStatement();
                
                 st1 = connect.createStatement();
                st2 = connect.createStatement();
                st3 = connect.createStatement();
                st4 = connect.createStatement();
                st0 = connect.createStatement();
                st = connect.createStatement();
                
                
 state=(Statement)connect.createStatement();
 state1=(Statement)connect.createStatement();
 state2=(Statement)connect.createStatement();
 state3=(Statement)connect.createStatement();
 state4=(Statement)connect.createStatement();


            }


        } catch (Exception ex) {
            Logger.getLogger(dbConn.class.getName()).log(Level.SEVERE, null, ex);
            System.out.println("ERROR WHILE CONNECTING TO DATABASE. CHECK YOUR CONNECTION CREDENTIALS SETTINGS in dbConn.java "+ex);
            //error in dbase configuration 
            //call the jsp page that does configuration

            if (issetdbcalled_exception%2 == 0) {
                calldbjsp();
                issetdbcalled_exception ++;
            }
            else{
            issetdbcalled_exception ++;
            }

        }
    }

    public final boolean getdbsettings(String drive) {
        boolean worked = true;

        try {



            String dbconnpath = drive + ":\\HALISI\\DATA\\DO_NOT_DELETE\\_\\_\\_\\dbconnection.txt";

            //File file = new File("");
            // InputStream inStream = getClass().getResourceAsStream("Web-INF/classes/dbconnection.txt");  
            FileInputStream fstream = new FileInputStream(dbconnpath);
            // Get the object of DataInputStream
            DataInputStream in = new DataInputStream(fstream);
            BufferedReader br = new BufferedReader(new InputStreamReader(in));
            String stLine;
            //Read File Line By Line
            int count = 0;
            while ((stLine = br.readLine()) != null) {

                // Print the content on the console


                dbsetup[count] = stLine;


                if (count < 4) {
                    count++;
                }
            }
            //Close the input stream
            in.close();
        } catch (IOException ex) {
            Logger.getLogger(dbConn.class.getName()).log(Level.SEVERE, null, ex);


            System.out.println("MY VALUE:" + issetdbcalled_file_exists);

            if (issetdbcalled_file_exists%2 == 0) {
                calldbjsp();
                issetdbcalled_file_exists++;
            }
            else{
            issetdbcalled_file_exists++;
            }

            System.out.println("MY VALUE:" + issetdbcalled_file_exists);


            System.out.println("ERROR:      FILE NOT FOUND");
            worked = false;

        }

        return worked;

    }

    public void calldbjsp() {
        try {

            //not so good for now because the host name is static


            String url = "http://localhost:8080/import711.jsp";
            java.awt.Desktop.getDesktop().browse(java.net.URI.create(url));
            //getdbsettings("M");
        } catch (IOException ex) {
            Logger.getLogger(dbConn.class.getName()).log(Level.SEVERE, null, ex);
        }


    }
}
