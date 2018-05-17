/*
 Notes: This raw data is for positive EID. The data doesnt have age and sex
 Age and sex should be gotten from the eid tested raw data during the importing of the raw data positives into the eid_datim_output table.

 */
package loaddata;

import db.dbConn;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.sql.SQLException;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.servlet.ServletException;
import javax.servlet.annotation.MultipartConfig;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.HttpSession;
import javax.servlet.http.Part;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

@MultipartConfig(fileSizeThreshold = 1024 * 1024 * 20, // 20 MB 
        maxFileSize = 1024 * 1024 * 50, // 50 MB
        maxRequestSize = 1024 * 1024 * 100)

/**
 *
 * @author Emmanuel Kaunda
 */
public class importdata extends HttpServlet {

    String full_path = "";
    String fileName = "";
    int checker_dist, checker_hf;
    File file_source;
    HttpSession session;
    private static final long serialVersionUID = 205242440643911308L;
    private static final String UPLOAD_DIR = "uploads";
    String nextpage = "";
    String quarterName, facilityName, facilityID, id, missingFacility;
    int nomflcodesites;

    String FPProgestinN;
    String FPProgestinR;
    String FPProgestinT;
    String FPCocN;
    String FPCocR;
    String FPCocT;
    String FPEcpN;
    String FPEcpR;
    String FPEcpT;
    String FPINJECTABLESN;
    String FPINJECTABLESR;
    String FPINJECTABLEST;
    String FPINJECTIONSN;
    String FPINJECTIONSR;
    String FPINJECTIONST;
    String FPIUCDN;
    String FPIUCDR;
    String FPIUCDT;
    String FPIMPLANTSN;
    String FPIMPLANTSR;
    String FPIMPLANTST;
    String FPBTLN;
    String FPBTLR;
    String FPBTLT;
    String FPVasectomyN;
    String FPVasectomyR;
    String FPVasectomyT;
    String FPCONDOMSMN;
    String FPCONDOMSFN;
    String FPCONDOMST;
    String FPNaturalN;
    String FPNaturalR;
    String FPNaturalT;
    String FPCLIENTSN;
    String FPCLIENTSR;
    String FPCLIENTST;
    String FPADOLESCENT10_14N;
    String FPADOLESCENT10_14R;
    String FPADOLESCENT10_14T;
    String FPADOLESCENT15_19N;
    String FPADOLESCENT15_19R;
    String FPADOLESCENT15_19T;
    String FPADOLESCENT20_24N;
    String FPADOLESCENT20_24R;
    String FPADOLESCENT20_24T;
    String FPIUCDRemoval;
    String FPIMPLANTSRemoval;
    String PMCTA_1stVisit_ANC;
    String PMCTA_ReVisit_ANC;
    String PMCTANCClientsT;
    String PMCTIPT1;
    String PMCTIPT2;
    String PMCTHB11;
    String PMCTANCClients4;
    String PMCTITN1;
    String PMCTITN;
    String PMTCTSYPHILISTES;
    String PMTCTSYPHILISPOS;
    String PMTCTCOUNSELLEDFEED;
    String PMTCTBREAST;
    String PMTCTEXERCISE;
    String PMTCTPREG10_14;
    String PMTCTPREG15_19;
    String PMTCTIRON;
    String PMTCTFOLIC;
    String PMTCTFERROUS;
    String MATNormalDelivery;
    String MATCSection;
    String MATBreech;
    String MATAssistedVag;
    String MATDeliveryT;
    String MATLiveBirth;
    String MATFreshStillBirth;
    String MATMeceratedStillBirth;
    String MATDeformities;
    String MATLowAPGAR;
    String MATWeight2500;
    String MATTetracycline;
    String MATPreTerm;
    String MATDischargealive;
    String MATbreastfeeding1;
    String MATDeliveriesPos;
    String MATNeoNatalD;
    String MATMaternalD10_19;
    String MATMaternalD;
    String MATMaternalDAudited;
    String MATAPHAlive;
    String MATAPHDead;
    String MATPPHAlive;
    String MATPPHDead;
    String MATEclampAlive;
    String MATEclampDead;
    String MATRupUtAlive;
    String MATRupUtDead;
    String MATObstrLaborAlive;
    String MATObstrLaborDead;
    String MATSepsisAlive;
    String MATSepsisDead;
    String MATREFFromOtherFacility;
    String MATREFFromCU;
    String MATREFToOtherFacility;
    String MATREFToCU;
    String SGBVRape72_0_9;
    String SGBVRape72_10_17;
    String SGBVRape72_18_49;
    String SGBVRape72_50;
    String SGBVRape72T;
    String SGBVinitPEP0_9;
    String SGBVinitPEP10_17;
    String SGBVinitPEP18_49;
    String SGBVinitPEP50;
    String SGBVinitPEPT;
    String SGBVcompPEP0_9;
    String SGBVcompPEP10_17;
    String SGBVcompPEP18_49;
    String SGBVcompPEP50;
    String SGBVcompPEPT;
    String SGBVPregnant0_9;
    String SGBVPregnant10_17;
    String SGBVPregnant18_49;
    String SGBVPregnant50;
    String SGBVPregnantT;
    String SGBVseroconverting0_9;
    String SGBVseroconverting10_17;
    String SGBVseroconverting18_49;
    String SGBVseroconverting50;
    String SGBVseroconvertingT;
    String SGBVsurvivors0_9;
    String SGBVsurvivors10_17;
    String SGBVsurvivors18_49;
    String SGBVsurvivors50;
    String SGBVsurvivorsT;
    String PAC10_19;
    String PACT;
    String CHANIS0_5NormalweightF;
    String CHANIS0_5NormalweightM;
    String CHANIS0_5NormalweightT;
    String CHANIS0_5UnderweightF;
    String CHANIS0_5UnderweightM;
    String CHANIS0_5UnderweightT;
    String CHANIS0_5sevUnderweightF;
    String CHANIS0_5sevUnderweightM;
    String CHANIS0_5sevUnderweightT;
    String CHANIS0_5OverweightF;
    String CHANIS0_5OverweightM;
    String CHANIS0_5OverweightT;
    String CHANIS0_5ObeseF;
    String CHANIS0_5ObeseM;
    String CHANIS0_5ObeseT;
    String CHANIS0_5TWF;
    String CHANIS0_5TWM;
    String CHANIS0_5TW;
    String CHANIS6_23NormalweightF;
    String CHANIS6_23NormalweightM;
    String CHANIS6_23NormalweightT;
    String CHANIS6_23UnderweightF;
    String CHANIS6_23UnderweightM;
    String CHANIS6_23UnderweightT;
    String CHANIS6_23sevUnderweightF;
    String CHANIS6_23sevUnderweightM;
    String CHANIS6_23sevUnderweightT;
    String CHANIS6_23OverweightF;
    String CHANIS6_23OverweightM;
    String CHANIS6_23OverweightT;
    String CHANIS6_23ObeseF;
    String CHANIS6_23ObeseM;
    String CHANIS6_23ObeseT;
    String CHANIS6_23TWF;
    String CHANIS6_23TWM;
    String CHANIS6_23TW;
    String CHANIS24_59NormalweightF;
    String CHANIS24_59NormalweightM;
    String CHANIS24_59NormalweightT;
    String CHANIS24_59UnderweightF;
    String CHANIS24_59UnderweightM;
    String CHANIS24_59UnderweightT;
    String CHANIS24_59sevUnderweightF;
    String CHANIS24_59sevUnderweightM;
    String CHANIS24_59sevUnderweightT;
    String CHANIS24_59OverweightF;
    String CHANIS24_59OverweightM;
    String CHANIS24_59OverweightT;
    String CHANIS24_59ObeseF;
    String CHANIS24_59ObeseM;
    String CHANIS24_59ObeseT;
    String CHANIS24_59TWF;
    String CHANIS24_59TWM;
    String CHANIS24_59TW;
    String CHANISMUACNormalF;
    String CHANISMUACNormalM;
    String CHANISMUACNormalT;
    String CHANISMUACModerateF;
    String CHANISMUACModerateM;
    String CHANISMUACModerateT;
    String CHANISMUACSevereF;
    String CHANISMUACSevereM;
    String CHANISMUACSevereT;
    String CHANISMUACMeasuredF;
    String CHANISMUACMeasuredM;
    String CHANISMUACMeasuredT;
    String CHANIS0_5NormalHeightF;
    String CHANIS0_5NormalHeightM;
    String CHANIS0_5NormalHeightT;
    String CHANIS0_5StuntedF;
    String CHANIS0_5StuntedM;
    String CHANIS0_5StuntedT;
    String CHANIS0_5sevStuntedF;
    String CHANIS0_5sevStuntedM;
    String CHANIS0_5sevStuntedT;
    String CHANIS0_5TMeasF;
    String CHANIS0_5TMeasM;
    String CHANIS0_5TMeas;
    String CHANIS6_23NormalHeightF;
    String CHANIS6_23NormalHeightM;
    String CHANIS6_23NormalHeightT;
    String CHANIS6_23StuntedF;
    String CHANIS6_23StuntedM;
    String CHANIS6_23StuntedT;
    String CHANIS6_23sevStuntedF;
    String CHANIS6_23sevStuntedM;
    String CHANIS6_23sevStuntedT;
    String CHANIS6_23TMeasF;
    String CHANIS6_23TMeasM;
    String CHANIS6_23TMeas;
    String CHANIS24_59NormalHeightF;
    String CHANIS24_59NormalHeightM;
    String CHANIS24_59NormalHeightT;
    String CHANIS24_59StuntedF;
    String CHANIS24_59StuntedM;
    String CHANIS24_59StuntedT;
    String CHANIS24_59sevStuntedF;
    String CHANIS24_59sevStuntedM;
    String CHANIS24_59sevStuntedT;
    String CHANIS24_59TMeasF;
    String CHANIS24_59TMeasM;
    String CHANIS24_59TMeas;
    String CHANIS0_59NewVisitsF;
    String CHANIS0_59NewVisitsM;
    String CHANIS0_59NewVisitsT;
    String CHANIS0_59KwashiakorF;
    String CHANIS0_59KwashiakorM;
    String CHANIS0_59KwashiakorT;
    String CHANIS0_59MarasmusF;
    String CHANIS0_59MarasmusM;
    String CHANIS0_59MarasmusT;
    String CHANIS0_59FalgrowthF;
    String CHANIS0_59FalgrowthM;
    String CHANIS0_59FalgrowthT;
    String CHANIS0_59F;
    String CHANIS0_59M;
    String CHANIS0_59T;
    String CHANIS0_5EXCLBreastF;
    String CHANIS0_5EXCLBreastM;
    String CHANIS0_5EXCLBreastT;
    String CHANIS12_59DewormedF;
    String CHANIS12_59DewormedM;
    String CHANIS12_59DewormedT;
    String CHANIS6_23MNPsF;
    String CHANIS6_23MNPsM;
    String CHANIS6_23MNPsT;
    String CHANIS0_59DisabilityF;
    String CHANIS0_59DisabilityM;
    String CHANIS0_59DisabilityT;
    String CCSVVH24;
    String CCSVVH25_49;
    String CCSVVH50;
    String CCSPAPSMEAR24;
    String CCSPAPSMEAR25_49;
    String CCSPAPSMEAR50;
    String CCSHPV24;
    String CCSHPV25_49;
    String CCSHPV50;
    String CCSVIAVILIPOS24;
    String CCSVIAVILIPOS25_49;
    String CCSVIAVILIPOS50;
    String CCSCYTOLPOS24;
    String CCSCYTOLPOS25_49;
    String CCSCYTOLPOS50;
    String CCSHPVPOS24;
    String CCSHPVPOS25_49;
    String CCSHPVPOS50;
    String CCSSUSPICIOUSLES24;
    String CCSSUSPICIOUSLES25_49;
    String CCSSUSPICIOUSLES50;
    String CCSCryotherapy24;
    String CCSCryotherapy25_49;
    String CCSCryotherapy50;
    String CCSLEEP24;
    String CCSLEEP25_49;
    String CCSLEEP50;
    String CCSHIVPOSSCREENED24;
    String CCSHIVPOSSCREENED25_49;
    String CCSHIVPOSSCREENED50;
    String PNCBreastExam;
    String PNCCounselled;
    String PNCFistula;
    String PNCExerNegative;
    String PNCExerPositive;
    String PNCCCSsuspect;
    String PNCmotherspostpartum2_3;
    String PNCmotherspostpartum6;
    String PNCinfantspostpartum2_3;
    String PNCinfantspostpartum6;
    String PNCreferralsfromotherHF;
    String PNCreferralsfromotherCU;
    String PNCreferralsTootherHF;
    String PNCreferralsTootherCU;
    String RsAssessed;
    String Rstreated;
    String RsRehabilitated;
    String Rsreffered;
    String RsIntergrated;
    String MSWpscounselling;
    String MSWdrugabuse;
    String MSWMental;
    String MSWAdolescent;
    String MSWPsAsses;
    String MSWsocialinv;
    String MSWsocialRehab;
    String MSWoutreach;
    String MSWreferrals;
    String MSWwaivedpatients;
    String PsPWDOPD4;
    String PsPWDOPD5_19;
    String PsPWDOPD20;
    String PsPWDinpatient4;
    String PsPWDinpatient5_19;
    String PsPWDinpatient20;
    String PsotherOPD4;
    String PsotherOPD5_19;
    String PsotherOPD20;
    String Psotherinpatient4;
    String Psotherinpatient5_19;
    String Psotherinpatient20;
    String PsTreatments4;
    String PsTreatments5_19;
    String PsTreatments20;
    String PsAssessed4;
    String PsAssessed5_19;
    String PsAssessed20;
    String PsServices4;
    String PsServices5_19;
    String PsServices20;
    String PsANCCounsel5_19;
    String PsANCCounsel20;
    String PsExercise5_19;
    String PsExercise20;
    String PsFIFcollected5_19;
    String PsFIFcollected20;
    String PsFIFwaived5_19;
    String PsFIFwaived20;
    String PsFIFexempted4;
    String PsFIFexempted5_19;
    String PsFIFexempted20;
    String PsDiasbilitymeeting4;
    String PsDiasbilitymeeting5_19;
    String PsDiasbilitymeeting20;

    String opd_partographs;
    String opd_oxytocyn;
    String opd_resucitated;

    String bcg_u1;
    String bcg_a1;
    String opv_w2wk;
    String opv1_u1;
    String opv1_a1;
    String opv2_u1;
    String opv2_a1;
    String opv3_u1;
    String opv3_a1;
    String ipv_u1;
    String ipv_a1;
    String dhh1_u1;
    String dhh1_a1;
    String dhh2_u1;
    String dhh2_a1;
    String dhh3_u1;
    String dhh3_a1;
    String pneumo1_u1;
    String pneumo1_a1;
    String pneumo2_u1;
    String pneumo2_a1;
    String pneumo3_u1;
    String pneumo3_a1;
    String rota1_u1;
    String rota2_u1;
    String vita_6;
    String yv_u1;
    String yv_a1;
    String mr1_u1;
    String mr1_a1;
    String fic_1;
    String vita_1yr;
    String vita_1half;
    String mr2_1half;
    String mr2_a2;
    String ttp_dose1;
    String ttp_dose2;
    String ttp_dose3;
    String ttp_dose4;
    String ttp_dose5;
    String ae_immun;
    String vita_2_5;
    String vita_lac_m;
    String squint_u1;
    String cce_type;
    String cce_model;
    String cce_sn;
    String cce_ws;
    String cce_es;
    String cce_age;
    String vac_type1;
    String vac_days1;
    String vac_type2;
    String vac_days2;
    String vita_type;
    String vita_days;
    String diarrhoea;
    String ors_zinc;
    String amoxycilin;
    String lastimportedsheet;

    int year, quarter, checker, missing = 0, added = 0, updated = 0,skipped=0;
    String county_name, county_id, district_name, district_id, hf_name, hf_id;

    String indicator = "";
    String reportingyear = "";
    String reportingmonth = "";
    String yearmonth = "";

    @Override
    protected void doPost(HttpServletRequest request, HttpServletResponse response)
            throws ServletException, IOException {

        missingFacility = "";
        FPProgestinN = "0";
        FPProgestinR = "0";
        FPProgestinT = "0";
        FPCocN = "0";
        FPCocR = "0";
        FPCocT = "0";
        FPEcpN = "0";
        FPEcpR = "0";
        FPEcpT = "0";
        FPINJECTABLESN = "0";
        FPINJECTABLESR = "0";
        FPINJECTABLEST = "0";
        FPINJECTIONSN = "0";
        FPINJECTIONSR = "0";
        FPINJECTIONST = "0";
        FPIUCDN = "0";
        FPIUCDR = "0";
        FPIUCDT = "0";
        FPIMPLANTSN = "0";
        FPIMPLANTSR = "0";
        FPIMPLANTST = "0";
        FPBTLN = "0";
        FPBTLR = "0";
        FPBTLT = "0";
        FPVasectomyN = "0";
        FPVasectomyR = "0";
        FPVasectomyT = "0";
        FPCONDOMSMN = "0";
        FPCONDOMSFN = "0";
        FPCONDOMST = "0";
        FPNaturalN = "0";
        FPNaturalR = "0";
        FPNaturalT = "0";
        FPCLIENTSN = "0";
        FPCLIENTSR = "0";
        FPCLIENTST = "0";
        FPADOLESCENT10_14N = "0";
        FPADOLESCENT10_14R = "0";
        FPADOLESCENT10_14T = "0";
        FPADOLESCENT15_19N = "0";
        FPADOLESCENT15_19R = "0";
        FPADOLESCENT15_19T = "0";
        FPADOLESCENT20_24N = "0";
        FPADOLESCENT20_24R = "0";
        FPADOLESCENT20_24T = "0";
        FPIUCDRemoval = "0";
        FPIMPLANTSRemoval = "0";
        PMCTA_1stVisit_ANC = "0";
        PMCTA_ReVisit_ANC = "0";
        PMCTANCClientsT = "0";
        PMCTIPT1 = "0";
        PMCTIPT2 = "0";
        PMCTHB11 = "0";
        PMCTANCClients4 = "0";
        PMCTITN1 = "0";
        PMCTITN = "0";
        PMTCTSYPHILISTES = "0";
        PMTCTSYPHILISPOS = "0";
        PMTCTCOUNSELLEDFEED = "0";
        PMTCTBREAST = "0";
        PMTCTEXERCISE = "0";
        PMTCTPREG10_14 = "0";
        PMTCTPREG15_19 = "0";
        PMTCTIRON = "0";
        PMTCTFOLIC = "0";
        PMTCTFERROUS = "0";
        MATNormalDelivery = "0";
        MATCSection = "0";
        MATBreech = "0";
        MATAssistedVag = "0";
        MATDeliveryT = "0";
        MATLiveBirth = "0";
        MATFreshStillBirth = "0";
        MATMeceratedStillBirth = "0";
        MATDeformities = "0";
        MATLowAPGAR = "0";
        MATWeight2500 = "0";
        MATTetracycline = "0";
        MATPreTerm = "0";
        MATDischargealive = "0";
        MATbreastfeeding1 = "0";
        MATDeliveriesPos = "0";
        MATNeoNatalD = "0";
        MATMaternalD10_19 = "0";
        MATMaternalD = "0";
        MATMaternalDAudited = "0";
        MATAPHAlive = "0";
        MATAPHDead = "0";
        MATPPHAlive = "0";
        MATPPHDead = "0";
        MATEclampAlive = "0";
        MATEclampDead = "0";
        MATRupUtAlive = "0";
        MATRupUtDead = "0";
        MATObstrLaborAlive = "0";
        MATObstrLaborDead = "0";
        MATSepsisAlive = "0";
        MATSepsisDead = "0";
        MATREFFromOtherFacility = "0";
        MATREFFromCU = "0";
        MATREFToOtherFacility = "0";
        MATREFToCU = "0";
        SGBVRape72_0_9 = "0";
        SGBVRape72_10_17 = "0";
        SGBVRape72_18_49 = "0";
        SGBVRape72_50 = "0";
        SGBVRape72T = "0";
        SGBVinitPEP0_9 = "0";
        SGBVinitPEP10_17 = "0";
        SGBVinitPEP18_49 = "0";
        SGBVinitPEP50 = "0";
        SGBVinitPEPT = "0";
        SGBVcompPEP0_9 = "0";
        SGBVcompPEP10_17 = "0";
        SGBVcompPEP18_49 = "0";
        SGBVcompPEP50 = "0";
        SGBVcompPEPT = "0";
        SGBVPregnant0_9 = "0";
        SGBVPregnant10_17 = "0";
        SGBVPregnant18_49 = "0";
        SGBVPregnant50 = "0";
        SGBVPregnantT = "0";
        SGBVseroconverting0_9 = "0";
        SGBVseroconverting10_17 = "0";
        SGBVseroconverting18_49 = "0";
        SGBVseroconverting50 = "0";
        SGBVseroconvertingT = "0";
        SGBVsurvivors0_9 = "0";
        SGBVsurvivors10_17 = "0";
        SGBVsurvivors18_49 = "0";
        SGBVsurvivors50 = "0";
        SGBVsurvivorsT = "0";
        PAC10_19 = "0";
        PACT = "0";
        CHANIS0_5NormalweightF = "0";
        CHANIS0_5NormalweightM = "0";
        CHANIS0_5NormalweightT = "0";
        CHANIS0_5UnderweightF = "0";
        CHANIS0_5UnderweightM = "0";
        CHANIS0_5UnderweightT = "0";
        CHANIS0_5sevUnderweightF = "0";
        CHANIS0_5sevUnderweightM = "0";
        CHANIS0_5sevUnderweightT = "0";
        CHANIS0_5OverweightF = "0";
        CHANIS0_5OverweightM = "0";
        CHANIS0_5OverweightT = "0";
        CHANIS0_5ObeseF = "0";
        CHANIS0_5ObeseM = "0";
        CHANIS0_5ObeseT = "0";
        CHANIS0_5TWF = "0";
        CHANIS0_5TWM = "0";
        CHANIS0_5TW = "0";
        CHANIS6_23NormalweightF = "0";
        CHANIS6_23NormalweightM = "0";
        CHANIS6_23NormalweightT = "0";
        CHANIS6_23UnderweightF = "0";
        CHANIS6_23UnderweightM = "0";
        CHANIS6_23UnderweightT = "0";
        CHANIS6_23sevUnderweightF = "0";
        CHANIS6_23sevUnderweightM = "0";
        CHANIS6_23sevUnderweightT = "0";
        CHANIS6_23OverweightF = "0";
        CHANIS6_23OverweightM = "0";
        CHANIS6_23OverweightT = "0";
        CHANIS6_23ObeseF = "0";
        CHANIS6_23ObeseM = "0";
        CHANIS6_23ObeseT = "0";
        CHANIS6_23TWF = "0";
        CHANIS6_23TWM = "0";
        CHANIS6_23TW = "0";
        CHANIS24_59NormalweightF = "0";
        CHANIS24_59NormalweightM = "0";
        CHANIS24_59NormalweightT = "0";
        CHANIS24_59UnderweightF = "0";
        CHANIS24_59UnderweightM = "0";
        CHANIS24_59UnderweightT = "0";
        CHANIS24_59sevUnderweightF = "0";
        CHANIS24_59sevUnderweightM = "0";
        CHANIS24_59sevUnderweightT = "0";
        CHANIS24_59OverweightF = "0";
        CHANIS24_59OverweightM = "0";
        CHANIS24_59OverweightT = "0";
        CHANIS24_59ObeseF = "0";
        CHANIS24_59ObeseM = "0";
        CHANIS24_59ObeseT = "0";
        CHANIS24_59TWF = "0";
        CHANIS24_59TWM = "0";
        CHANIS24_59TW = "0";
        CHANISMUACNormalF = "0";
        CHANISMUACNormalM = "0";
        CHANISMUACNormalT = "0";
        CHANISMUACModerateF = "0";
        CHANISMUACModerateM = "0";
        CHANISMUACModerateT = "0";
        CHANISMUACSevereF = "0";
        CHANISMUACSevereM = "0";
        CHANISMUACSevereT = "0";
        CHANISMUACMeasuredF = "0";
        CHANISMUACMeasuredM = "0";
        CHANISMUACMeasuredT = "0";
        CHANIS0_5NormalHeightF = "0";
        CHANIS0_5NormalHeightM = "0";
        CHANIS0_5NormalHeightT = "0";
        CHANIS0_5StuntedF = "0";
        CHANIS0_5StuntedM = "0";
        CHANIS0_5StuntedT = "0";
        CHANIS0_5sevStuntedF = "0";
        CHANIS0_5sevStuntedM = "0";
        CHANIS0_5sevStuntedT = "0";
        CHANIS0_5TMeasF = "0";
        CHANIS0_5TMeasM = "0";
        CHANIS0_5TMeas = "0";
        CHANIS6_23NormalHeightF = "0";
        CHANIS6_23NormalHeightM = "0";
        CHANIS6_23NormalHeightT = "0";
        CHANIS6_23StuntedF = "0";
        CHANIS6_23StuntedM = "0";
        CHANIS6_23StuntedT = "0";
        CHANIS6_23sevStuntedF = "0";
        CHANIS6_23sevStuntedM = "0";
        CHANIS6_23sevStuntedT = "0";
        CHANIS6_23TMeasF = "0";
        CHANIS6_23TMeasM = "0";
        CHANIS6_23TMeas = "0";
        CHANIS24_59NormalHeightF = "0";
        CHANIS24_59NormalHeightM = "0";
        CHANIS24_59NormalHeightT = "0";
        CHANIS24_59StuntedF = "0";
        CHANIS24_59StuntedM = "0";
        CHANIS24_59StuntedT = "0";
        CHANIS24_59sevStuntedF = "0";
        CHANIS24_59sevStuntedM = "0";
        CHANIS24_59sevStuntedT = "0";
        CHANIS24_59TMeasF = "0";
        CHANIS24_59TMeasM = "0";
        CHANIS24_59TMeas = "0";
        CHANIS0_59NewVisitsF = "0";
        CHANIS0_59NewVisitsM = "0";
        CHANIS0_59NewVisitsT = "0";
        CHANIS0_59KwashiakorF = "0";
        CHANIS0_59KwashiakorM = "0";
        CHANIS0_59KwashiakorT = "0";
        CHANIS0_59MarasmusF = "0";
        CHANIS0_59MarasmusM = "0";
        CHANIS0_59MarasmusT = "0";
        CHANIS0_59FalgrowthF = "0";
        CHANIS0_59FalgrowthM = "0";
        CHANIS0_59FalgrowthT = "0";
        CHANIS0_59F = "0";
        CHANIS0_59M = "0";
        CHANIS0_59T = "0";
        CHANIS0_5EXCLBreastF = "0";
        CHANIS0_5EXCLBreastM = "0";
        CHANIS0_5EXCLBreastT = "0";
        CHANIS12_59DewormedF = "0";
        CHANIS12_59DewormedM = "0";
        CHANIS12_59DewormedT = "0";
        CHANIS6_23MNPsF = "0";
        CHANIS6_23MNPsM = "0";
        CHANIS6_23MNPsT = "0";
        CHANIS0_59DisabilityF = "0";
        CHANIS0_59DisabilityM = "0";
        CHANIS0_59DisabilityT = "0";
        CCSVVH24 = "0";
        CCSVVH25_49 = "0";
        CCSVVH50 = "0";
        CCSPAPSMEAR24 = "0";
        CCSPAPSMEAR25_49 = "0";
        CCSPAPSMEAR50 = "0";
        CCSHPV24 = "0";
        CCSHPV25_49 = "0";
        CCSHPV50 = "0";
        CCSVIAVILIPOS24 = "0";
        CCSVIAVILIPOS25_49 = "0";
        CCSVIAVILIPOS50 = "0";
        CCSCYTOLPOS24 = "0";
        CCSCYTOLPOS25_49 = "0";
        CCSCYTOLPOS50 = "0";
        CCSHPVPOS24 = "0";
        CCSHPVPOS25_49 = "0";
        CCSHPVPOS50 = "0";
        CCSSUSPICIOUSLES24 = "0";
        CCSSUSPICIOUSLES25_49 = "0";
        CCSSUSPICIOUSLES50 = "0";
        CCSCryotherapy24 = "0";
        CCSCryotherapy25_49 = "0";
        CCSCryotherapy50 = "0";
        CCSLEEP24 = "0";
        CCSLEEP25_49 = "0";
        CCSLEEP50 = "0";
        CCSHIVPOSSCREENED24 = "0";
        CCSHIVPOSSCREENED25_49 = "0";
        CCSHIVPOSSCREENED50 = "0";
        PNCBreastExam = "0";
        PNCCounselled = "0";
        PNCFistula = "0";
        PNCExerNegative = "0";
        PNCExerPositive = "0";
        PNCCCSsuspect = "0";
        PNCmotherspostpartum2_3 = "0";
        PNCmotherspostpartum6 = "0";
        PNCinfantspostpartum2_3 = "0";
        PNCinfantspostpartum6 = "0";
        PNCreferralsfromotherHF = "0";
        PNCreferralsfromotherCU = "0";
        PNCreferralsTootherHF = "0";
        PNCreferralsTootherCU = "0";
        RsAssessed = "0";
        Rstreated = "0";
        RsRehabilitated = "0";
        Rsreffered = "0";
        RsIntergrated = "0";
        MSWpscounselling = "0";
        MSWdrugabuse = "0";
        MSWMental = "0";
        MSWAdolescent = "0";
        MSWPsAsses = "0";
        MSWsocialinv = "0";
        MSWsocialRehab = "0";
        MSWoutreach = "0";
        MSWreferrals = "0";
        MSWwaivedpatients = "0";
        PsPWDOPD4 = "0";
        PsPWDOPD5_19 = "0";
        PsPWDOPD20 = "0";
        PsPWDinpatient4 = "0";
        PsPWDinpatient5_19 = "0";
        PsPWDinpatient20 = "0";
        PsotherOPD4 = "0";
        PsotherOPD5_19 = "0";
        PsotherOPD20 = "0";
        Psotherinpatient4 = "0";
        Psotherinpatient5_19 = "0";
        Psotherinpatient20 = "0";
        PsTreatments4 = "0";
        PsTreatments5_19 = "0";
        PsTreatments20 = "0";
        PsAssessed4 = "0";
        PsAssessed5_19 = "0";
        PsAssessed20 = "0";
        PsServices4 = "0";
        PsServices5_19 = "0";
        PsServices20 = "0";
        PsANCCounsel5_19 = "0";
        PsANCCounsel20 = "0";
        PsExercise5_19 = "0";
        PsExercise20 = "0";
        PsFIFcollected5_19 = "0";
        PsFIFcollected20 = "0";
        PsFIFwaived5_19 = "0";
        PsFIFwaived20 = "0";
        PsFIFexempted4 = "0";
        PsFIFexempted5_19 = "0";
        PsFIFexempted20 = "0";
        PsDiasbilitymeeting4 = "0";
        PsDiasbilitymeeting5_19 = "0";
        PsDiasbilitymeeting20 = "0";

        opd_partographs = "0";
        opd_oxytocyn = "0";
        opd_resucitated = "0";

        bcg_u1 = "0";
        bcg_a1 = "0";
        opv_w2wk = "0";
        opv1_u1 = "0";
        opv1_a1 = "0";
        opv2_u1 = "0";
        opv2_a1 = "0";
        opv3_u1 = "0";
        opv3_a1 = "0";
        ipv_u1 = "0";
        ipv_a1 = "0";
        dhh1_u1 = "0";
        dhh1_a1 = "0";
        dhh2_u1 = "0";
        dhh2_a1 = "0";
        dhh3_u1 = "0";
        dhh3_a1 = "0";
        pneumo1_u1 = "0";
        pneumo1_a1 = "0";
        pneumo2_u1 = "0";
        pneumo2_a1 = "0";
        pneumo3_u1 = "0";
        pneumo3_a1 = "0";
        rota1_u1 = "0";
        rota2_u1 = "0";
        vita_6 = "0";
        yv_u1 = "0";
        yv_a1 = "0";
        mr1_u1 = "0";
        mr1_a1 = "0";
        fic_1 = "0";
        vita_1yr = "0";
        vita_1half = "0";
        mr2_1half = "0";
        mr2_a2 = "0";
        ttp_dose1 = "0";
        ttp_dose2 = "0";
        ttp_dose3 = "0";
        ttp_dose4 = "0";
        ttp_dose5 = "0";
        ae_immun = "0";
        vita_2_5 = "0";
        vita_lac_m = "0";
        squint_u1 = "0";
        cce_type = "0";
        cce_model = "0";
        cce_sn = "0";
        cce_ws = "0";
        cce_es = "0";
        cce_age = "0";
        vac_type1 = "0";
        vac_days1 = "0";
        vac_type2 = "0";
        vac_days2 = "0";
        vita_type = "0";
        vita_days = "0";
        diarrhoea = "0";
        ors_zinc = "0";
        amoxycilin = "0";

        missing = 0;
        added = 0;
        updated = 0;
        session = request.getSession();
        nomflcodesites = 0;

        /**
         * * *
         **
         */
        id = "";

        reportingmonth = request.getParameter("month");
        reportingyear = request.getParameter("year");

        if (!reportingyear.equals("")) {
            session.setAttribute("reportingyear", reportingyear);
            session.setAttribute("reportingmonth", reportingmonth);
        }
        String updatedfacil = "";
        String insertedfacil = "";
        String missingwithdatafacil = "";

        String mflcode = "";
        String serialnumber = "";
        String dbname = "baseline_checklist";

        dbConn conn = new dbConn();
        nextpage = "index.jsp";
        String applicationPath = request.getServletContext().getRealPath("");
        String uploadFilePath = applicationPath + File.separator + UPLOAD_DIR;
        session = request.getSession();
        File fileSaveDir = new File(uploadFilePath);
        if (!fileSaveDir.exists()) {
            fileSaveDir.mkdirs();
        }
        System.out.println("Upload File Directory=" + fileSaveDir.getAbsolutePath());

        for (Part part : request.getParts()) {
            if (!getFileName(part).equals("")) {
                fileName = getFileName(part);
                part.write(uploadFilePath + File.separator + fileName);
            }
        }

        if (!fileName.endsWith(".xlsm") && !fileName.endsWith(".xlsx")) {
            nextpage = "index.jsp";
            session.setAttribute("upload_success", "<font color=\"red\">Failed to load the excel file. Please choose a .xlsx excel file .</font>");
        } else {

            full_path = fileSaveDir.getAbsolutePath() + "\\" + fileName;

            System.out.println("the saved file directory is  :  " + full_path);
// GET DATA FROM THE EXCEL AND AND OUTPUT IT ON THE CONSOLE..................................

            FileInputStream fileInputStream = new FileInputStream(full_path);
            XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);

            int totalsheets = workbook.getNumberOfSheets();
            DataFormatter formatter = new DataFormatter(); //creating formatter using the default locale
            lastimportedsheet = "";
added=0;
skipped=0;
missing=0;
            for (int a = 0; a < totalsheets; a++) {

                XSSFSheet worksheet = workbook.getSheetAt(a);

                System.out.println(a + " (" + workbook.getSheetName(a) + ") out of " + totalsheets + " sheets");

                String sheetname = workbook.getSheetName(a);
                lastimportedsheet = (a+1) + " (" + workbook.getSheetName(a) + ") out of " + totalsheets + " sheets";
//_______        
                //if(1==1){
                if (sheetname.equals("Interface") || sheetname.equals("Login")) {

                } else if (worksheet.getRow(4) != null) {

                    int row = 4, col = 0;
                        //static number of rows

                    XSSFRow rowi = worksheet.getRow(row);
                    //-----------mfl-----------------------
                    XSSFCell cellfacil = worksheet.getRow(4).getCell((short) 5);
                    if (cellfacil.getCellType() == 0) {  //numeric
                        facilityName = "" + (int) cellfacil.getNumericCellValue();
                    } else if (cellfacil.getCellType() == 1) {
                        facilityName = cellfacil.getStringCellValue();
                    } else if (cellfacil.getCellType() == 2) {
                        facilityName = cellfacil.getRawValue();

                    } else {
                        facilityName = cellfacil.getRawValue();
                    }
                    //-----------mfl-----------------------
                    XSSFCell cellym = worksheet.getRow(4).getCell((short) 14);
                    if (cellym.getCellType() == 0) {  //numeric
                        yearmonth = "" + (int) cellym.getNumericCellValue();
                    } else if (cellym.getCellType() == 1) {
                        yearmonth = cellym.getStringCellValue();
                    } else if (cellym.getCellType() == 2) {
                        yearmonth = cellym.getRawValue();

                    } else {
                        yearmonth = cellym.getRawValue();
                    }

                    yearmonth = reportingyear + "" + reportingmonth;

                   XSSFCell cl_FPProgestinN = worksheet.getRow(35).getCell((short) 8); if ( cl_FPProgestinN!=null){ if ( cl_FPProgestinN.getCellType()==0) {FPProgestinN = "" + (int) cl_FPProgestinN.getNumericCellValue();} else if (cl_FPProgestinN.getCellType()==1) {FPProgestinN=cl_FPProgestinN.getStringCellValue();}  else if (cl_FPProgestinN.getCellType()== 2) {FPProgestinN=cl_FPProgestinN.getRawValue();  } else { FPProgestinN = "0";}}
XSSFCell cl_FPProgestinR = worksheet.getRow(35).getCell((short) 9); if ( cl_FPProgestinR!=null){ if ( cl_FPProgestinR.getCellType()==0) {FPProgestinR = "" + (int) cl_FPProgestinR.getNumericCellValue();} else if (cl_FPProgestinR.getCellType()==1) {FPProgestinR=cl_FPProgestinR.getStringCellValue();}  else if (cl_FPProgestinR.getCellType()== 2) {FPProgestinR=cl_FPProgestinR.getRawValue();  } else { FPProgestinR = "0";}}
XSSFCell cl_FPCocN = worksheet.getRow(36).getCell((short) 8); if ( cl_FPCocN!=null){ if ( cl_FPCocN.getCellType()==0) {FPCocN = "" + (int) cl_FPCocN.getNumericCellValue();} else if (cl_FPCocN.getCellType()==1) {FPCocN=cl_FPCocN.getStringCellValue();}  else if (cl_FPCocN.getCellType()== 2) {FPCocN=cl_FPCocN.getRawValue();  } else { FPCocN = "0";}}
XSSFCell cl_FPCocR = worksheet.getRow(36).getCell((short) 9); if ( cl_FPCocR!=null){ if ( cl_FPCocR.getCellType()==0) {FPCocR = "" + (int) cl_FPCocR.getNumericCellValue();} else if (cl_FPCocR.getCellType()==1) {FPCocR=cl_FPCocR.getStringCellValue();}  else if (cl_FPCocR.getCellType()== 2) {FPCocR=cl_FPCocR.getRawValue();  } else { FPCocR = "0";}}
XSSFCell cl_FPEcpN = worksheet.getRow(37).getCell((short) 8); if ( cl_FPEcpN!=null){ if ( cl_FPEcpN.getCellType()==0) {FPEcpN = "" + (int) cl_FPEcpN.getNumericCellValue();} else if (cl_FPEcpN.getCellType()==1) {FPEcpN=cl_FPEcpN.getStringCellValue();}  else if (cl_FPEcpN.getCellType()== 2) {FPEcpN=cl_FPEcpN.getRawValue();  } else { FPEcpN = "0";}}
XSSFCell cl_FPEcpR = worksheet.getRow(37).getCell((short) 9); if ( cl_FPEcpR!=null){ if ( cl_FPEcpR.getCellType()==0) {FPEcpR = "" + (int) cl_FPEcpR.getNumericCellValue();} else if (cl_FPEcpR.getCellType()==1) {FPEcpR=cl_FPEcpR.getStringCellValue();}  else if (cl_FPEcpR.getCellType()== 2) {FPEcpR=cl_FPEcpR.getRawValue();  } else { FPEcpR = "0";}}
XSSFCell cl_FPINJECTABLESN = worksheet.getRow(39).getCell((short) 8); if ( cl_FPINJECTABLESN!=null){ if ( cl_FPINJECTABLESN.getCellType()==0) {FPINJECTABLESN = "" + (int) cl_FPINJECTABLESN.getNumericCellValue();} else if (cl_FPINJECTABLESN.getCellType()==1) {FPINJECTABLESN=cl_FPINJECTABLESN.getStringCellValue();}  else if (cl_FPINJECTABLESN.getCellType()== 2) {FPINJECTABLESN=cl_FPINJECTABLESN.getRawValue();  } else { FPINJECTABLESN = "0";}}
XSSFCell cl_FPINJECTABLESR = worksheet.getRow(39).getCell((short) 9); if ( cl_FPINJECTABLESR!=null){ if ( cl_FPINJECTABLESR.getCellType()==0) {FPINJECTABLESR = "" + (int) cl_FPINJECTABLESR.getNumericCellValue();} else if (cl_FPINJECTABLESR.getCellType()==1) {FPINJECTABLESR=cl_FPINJECTABLESR.getStringCellValue();}  else if (cl_FPINJECTABLESR.getCellType()== 2) {FPINJECTABLESR=cl_FPINJECTABLESR.getRawValue();  } else { FPINJECTABLESR = "0";}}
XSSFCell cl_FPINJECTIONSN = worksheet.getRow(39).getCell((short) 3); if ( cl_FPINJECTIONSN!=null){ if ( cl_FPINJECTIONSN.getCellType()==0) {FPINJECTIONSN = "" + (int) cl_FPINJECTIONSN.getNumericCellValue();} else if (cl_FPINJECTIONSN.getCellType()==1) {FPINJECTIONSN=cl_FPINJECTIONSN.getStringCellValue();}  else if (cl_FPINJECTIONSN.getCellType()== 2) {FPINJECTIONSN=cl_FPINJECTIONSN.getRawValue();  } else { FPINJECTIONSN = "0";}}
XSSFCell cl_FPINJECTIONSR = worksheet.getRow(39).getCell((short) 4); if ( cl_FPINJECTIONSR!=null){ if ( cl_FPINJECTIONSR.getCellType()==0) {FPINJECTIONSR = "" + (int) cl_FPINJECTIONSR.getNumericCellValue();} else if (cl_FPINJECTIONSR.getCellType()==1) {FPINJECTIONSR=cl_FPINJECTIONSR.getStringCellValue();}  else if (cl_FPINJECTIONSR.getCellType()== 2) {FPINJECTIONSR=cl_FPINJECTIONSR.getRawValue();  } else { FPINJECTIONSR = "0";}}
XSSFCell cl_FPIUCDN = worksheet.getRow(40).getCell((short) 3); if ( cl_FPIUCDN!=null){ if ( cl_FPIUCDN.getCellType()==0) {FPIUCDN = "" + (int) cl_FPIUCDN.getNumericCellValue();} else if (cl_FPIUCDN.getCellType()==1) {FPIUCDN=cl_FPIUCDN.getStringCellValue();}  else if (cl_FPIUCDN.getCellType()== 2) {FPIUCDN=cl_FPIUCDN.getRawValue();  } else { FPIUCDN = "0";}}
XSSFCell cl_FPIUCDR = worksheet.getRow(40).getCell((short) 4); if ( cl_FPIUCDR!=null){ if ( cl_FPIUCDR.getCellType()==0) {FPIUCDR = "" + (int) cl_FPIUCDR.getNumericCellValue();} else if (cl_FPIUCDR.getCellType()==1) {FPIUCDR=cl_FPIUCDR.getStringCellValue();}  else if (cl_FPIUCDR.getCellType()== 2) {FPIUCDR=cl_FPIUCDR.getRawValue();  } else { FPIUCDR = "0";}}
XSSFCell cl_FPIMPLANTSN = worksheet.getRow(41).getCell((short) 3); if ( cl_FPIMPLANTSN!=null){ if ( cl_FPIMPLANTSN.getCellType()==0) {FPIMPLANTSN = "" + (int) cl_FPIMPLANTSN.getNumericCellValue();} else if (cl_FPIMPLANTSN.getCellType()==1) {FPIMPLANTSN=cl_FPIMPLANTSN.getStringCellValue();}  else if (cl_FPIMPLANTSN.getCellType()== 2) {FPIMPLANTSN=cl_FPIMPLANTSN.getRawValue();  } else { FPIMPLANTSN = "0";}}
XSSFCell cl_FPIMPLANTSR = worksheet.getRow(41).getCell((short) 4); if ( cl_FPIMPLANTSR!=null){ if ( cl_FPIMPLANTSR.getCellType()==0) {FPIMPLANTSR = "" + (int) cl_FPIMPLANTSR.getNumericCellValue();} else if (cl_FPIMPLANTSR.getCellType()==1) {FPIMPLANTSR=cl_FPIMPLANTSR.getStringCellValue();}  else if (cl_FPIMPLANTSR.getCellType()== 2) {FPIMPLANTSR=cl_FPIMPLANTSR.getRawValue();  } else { FPIMPLANTSR = "0";}}
XSSFCell cl_FPBTLN = worksheet.getRow(41).getCell((short) 8); if ( cl_FPBTLN!=null){ if ( cl_FPBTLN.getCellType()==0) {FPBTLN = "" + (int) cl_FPBTLN.getNumericCellValue();} else if (cl_FPBTLN.getCellType()==1) {FPBTLN=cl_FPBTLN.getStringCellValue();}  else if (cl_FPBTLN.getCellType()== 2) {FPBTLN=cl_FPBTLN.getRawValue();  } else { FPBTLN = "0";}}
XSSFCell cl_FPBTLR = worksheet.getRow(41).getCell((short) 9); if ( cl_FPBTLR!=null){ if ( cl_FPBTLR.getCellType()==0) {FPBTLR = "" + (int) cl_FPBTLR.getNumericCellValue();} else if (cl_FPBTLR.getCellType()==1) {FPBTLR=cl_FPBTLR.getStringCellValue();}  else if (cl_FPBTLR.getCellType()== 2) {FPBTLR=cl_FPBTLR.getRawValue();  } else { FPBTLR = "0";}}
XSSFCell cl_FPVasectomyN = worksheet.getRow(42).getCell((short) 8); if ( cl_FPVasectomyN!=null){ if ( cl_FPVasectomyN.getCellType()==0) {FPVasectomyN = "" + (int) cl_FPVasectomyN.getNumericCellValue();} else if (cl_FPVasectomyN.getCellType()==1) {FPVasectomyN=cl_FPVasectomyN.getStringCellValue();}  else if (cl_FPVasectomyN.getCellType()== 2) {FPVasectomyN=cl_FPVasectomyN.getRawValue();  } else { FPVasectomyN = "0";}}
XSSFCell cl_FPVasectomyR = worksheet.getRow(42).getCell((short) 9); if ( cl_FPVasectomyR!=null){ if ( cl_FPVasectomyR.getCellType()==0) {FPVasectomyR = "" + (int) cl_FPVasectomyR.getNumericCellValue();} else if (cl_FPVasectomyR.getCellType()==1) {FPVasectomyR=cl_FPVasectomyR.getStringCellValue();}  else if (cl_FPVasectomyR.getCellType()== 2) {FPVasectomyR=cl_FPVasectomyR.getRawValue();  } else { FPVasectomyR = "0";}}
XSSFCell cl_FPCONDOMSMN = worksheet.getRow(43).getCell((short) 8); if ( cl_FPCONDOMSMN!=null){ if ( cl_FPCONDOMSMN.getCellType()==0) {FPCONDOMSMN = "" + (int) cl_FPCONDOMSMN.getNumericCellValue();} else if (cl_FPCONDOMSMN.getCellType()==1) {FPCONDOMSMN=cl_FPCONDOMSMN.getStringCellValue();}  else if (cl_FPCONDOMSMN.getCellType()== 2) {FPCONDOMSMN=cl_FPCONDOMSMN.getRawValue();  } else { FPCONDOMSMN = "0";}}
XSSFCell cl_FPCONDOMSFN = worksheet.getRow(44).getCell((short) 8); if ( cl_FPCONDOMSFN!=null){ if ( cl_FPCONDOMSFN.getCellType()==0) {FPCONDOMSFN = "" + (int) cl_FPCONDOMSFN.getNumericCellValue();} else if (cl_FPCONDOMSFN.getCellType()==1) {FPCONDOMSFN=cl_FPCONDOMSFN.getStringCellValue();}  else if (cl_FPCONDOMSFN.getCellType()== 2) {FPCONDOMSFN=cl_FPCONDOMSFN.getRawValue();  } else { FPCONDOMSFN = "0";}}
XSSFCell cl_FPNaturalN = worksheet.getRow(45).getCell((short) 8); if ( cl_FPNaturalN!=null){ if ( cl_FPNaturalN.getCellType()==0) {FPNaturalN = "" + (int) cl_FPNaturalN.getNumericCellValue();} else if (cl_FPNaturalN.getCellType()==1) {FPNaturalN=cl_FPNaturalN.getStringCellValue();}  else if (cl_FPNaturalN.getCellType()== 2) {FPNaturalN=cl_FPNaturalN.getRawValue();  } else { FPNaturalN = "0";}}
XSSFCell cl_FPNaturalR = worksheet.getRow(45).getCell((short) 9); if ( cl_FPNaturalR!=null){ if ( cl_FPNaturalR.getCellType()==0) {FPNaturalR = "" + (int) cl_FPNaturalR.getNumericCellValue();} else if (cl_FPNaturalR.getCellType()==1) {FPNaturalR=cl_FPNaturalR.getStringCellValue();}  else if (cl_FPNaturalR.getCellType()== 2) {FPNaturalR=cl_FPNaturalR.getRawValue();  } else { FPNaturalR = "0";}}
XSSFCell cl_FPCLIENTSN = worksheet.getRow(46).getCell((short) 8); if ( cl_FPCLIENTSN!=null){ if ( cl_FPCLIENTSN.getCellType()==0) {FPCLIENTSN = "" + (int) cl_FPCLIENTSN.getNumericCellValue();} else if (cl_FPCLIENTSN.getCellType()==1) {FPCLIENTSN=cl_FPCLIENTSN.getStringCellValue();}  else if (cl_FPCLIENTSN.getCellType()== 2) {FPCLIENTSN=cl_FPCLIENTSN.getRawValue();  } else { FPCLIENTSN = "0";}}
XSSFCell cl_FPCLIENTSR = worksheet.getRow(46).getCell((short) 9); if ( cl_FPCLIENTSR!=null){ if ( cl_FPCLIENTSR.getCellType()==0) {FPCLIENTSR = "" + (int) cl_FPCLIENTSR.getNumericCellValue();} else if (cl_FPCLIENTSR.getCellType()==1) {FPCLIENTSR=cl_FPCLIENTSR.getStringCellValue();}  else if (cl_FPCLIENTSR.getCellType()== 2) {FPCLIENTSR=cl_FPCLIENTSR.getRawValue();  } else { FPCLIENTSR = "0";}}
XSSFCell cl_FPADOLESCENT10_14N = worksheet.getRow(47).getCell((short) 8); if ( cl_FPADOLESCENT10_14N!=null){ if ( cl_FPADOLESCENT10_14N.getCellType()==0) {FPADOLESCENT10_14N = "" + (int) cl_FPADOLESCENT10_14N.getNumericCellValue();} else if (cl_FPADOLESCENT10_14N.getCellType()==1) {FPADOLESCENT10_14N=cl_FPADOLESCENT10_14N.getStringCellValue();}  else if (cl_FPADOLESCENT10_14N.getCellType()== 2) {FPADOLESCENT10_14N=cl_FPADOLESCENT10_14N.getRawValue();  } else { FPADOLESCENT10_14N = "0";}}
XSSFCell cl_FPADOLESCENT10_14R = worksheet.getRow(47).getCell((short) 9); if ( cl_FPADOLESCENT10_14R!=null){ if ( cl_FPADOLESCENT10_14R.getCellType()==0) {FPADOLESCENT10_14R = "" + (int) cl_FPADOLESCENT10_14R.getNumericCellValue();} else if (cl_FPADOLESCENT10_14R.getCellType()==1) {FPADOLESCENT10_14R=cl_FPADOLESCENT10_14R.getStringCellValue();}  else if (cl_FPADOLESCENT10_14R.getCellType()== 2) {FPADOLESCENT10_14R=cl_FPADOLESCENT10_14R.getRawValue();  } else { FPADOLESCENT10_14R = "0";}}
XSSFCell cl_FPADOLESCENT15_19N = worksheet.getRow(48).getCell((short) 8); if ( cl_FPADOLESCENT15_19N!=null){ if ( cl_FPADOLESCENT15_19N.getCellType()==0) {FPADOLESCENT15_19N = "" + (int) cl_FPADOLESCENT15_19N.getNumericCellValue();} else if (cl_FPADOLESCENT15_19N.getCellType()==1) {FPADOLESCENT15_19N=cl_FPADOLESCENT15_19N.getStringCellValue();}  else if (cl_FPADOLESCENT15_19N.getCellType()== 2) {FPADOLESCENT15_19N=cl_FPADOLESCENT15_19N.getRawValue();  } else { FPADOLESCENT15_19N = "0";}}
XSSFCell cl_FPADOLESCENT15_19R = worksheet.getRow(48).getCell((short) 9); if ( cl_FPADOLESCENT15_19R!=null){ if ( cl_FPADOLESCENT15_19R.getCellType()==0) {FPADOLESCENT15_19R = "" + (int) cl_FPADOLESCENT15_19R.getNumericCellValue();} else if (cl_FPADOLESCENT15_19R.getCellType()==1) {FPADOLESCENT15_19R=cl_FPADOLESCENT15_19R.getStringCellValue();}  else if (cl_FPADOLESCENT15_19R.getCellType()== 2) {FPADOLESCENT15_19R=cl_FPADOLESCENT15_19R.getRawValue();  } else { FPADOLESCENT15_19R = "0";}}
XSSFCell cl_FPADOLESCENT20_24N = worksheet.getRow(49).getCell((short) 8); if ( cl_FPADOLESCENT20_24N!=null){ if ( cl_FPADOLESCENT20_24N.getCellType()==0) {FPADOLESCENT20_24N = "" + (int) cl_FPADOLESCENT20_24N.getNumericCellValue();} else if (cl_FPADOLESCENT20_24N.getCellType()==1) {FPADOLESCENT20_24N=cl_FPADOLESCENT20_24N.getStringCellValue();}  else if (cl_FPADOLESCENT20_24N.getCellType()== 2) {FPADOLESCENT20_24N=cl_FPADOLESCENT20_24N.getRawValue();  } else { FPADOLESCENT20_24N = "0";}}
XSSFCell cl_FPADOLESCENT20_24R = worksheet.getRow(49).getCell((short) 9); if ( cl_FPADOLESCENT20_24R!=null){ if ( cl_FPADOLESCENT20_24R.getCellType()==0) {FPADOLESCENT20_24R = "" + (int) cl_FPADOLESCENT20_24R.getNumericCellValue();} else if (cl_FPADOLESCENT20_24R.getCellType()==1) {FPADOLESCENT20_24R=cl_FPADOLESCENT20_24R.getStringCellValue();}  else if (cl_FPADOLESCENT20_24R.getCellType()== 2) {FPADOLESCENT20_24R=cl_FPADOLESCENT20_24R.getRawValue();  } else { FPADOLESCENT20_24R = "0";}}
XSSFCell cl_FPIUCDRemoval = worksheet.getRow(50).getCell((short) 8); if ( cl_FPIUCDRemoval!=null){ if ( cl_FPIUCDRemoval.getCellType()==0) {FPIUCDRemoval = "" + (int) cl_FPIUCDRemoval.getNumericCellValue();} else if (cl_FPIUCDRemoval.getCellType()==1) {FPIUCDRemoval=cl_FPIUCDRemoval.getStringCellValue();}  else if (cl_FPIUCDRemoval.getCellType()== 2) {FPIUCDRemoval=cl_FPIUCDRemoval.getRawValue();  } else { FPIUCDRemoval = "0";}}
XSSFCell cl_FPIMPLANTSRemoval = worksheet.getRow(51).getCell((short) 8); if ( cl_FPIMPLANTSRemoval!=null){ if ( cl_FPIMPLANTSRemoval.getCellType()==0) {FPIMPLANTSRemoval = "" + (int) cl_FPIMPLANTSRemoval.getNumericCellValue();} else if (cl_FPIMPLANTSRemoval.getCellType()==1) {FPIMPLANTSRemoval=cl_FPIMPLANTSRemoval.getStringCellValue();}  else if (cl_FPIMPLANTSRemoval.getCellType()== 2) {FPIMPLANTSRemoval=cl_FPIMPLANTSRemoval.getRawValue();  } else { FPIMPLANTSRemoval = "0";}}
XSSFCell cl_PMCTA_1stVisit_ANC = worksheet.getRow(7).getCell((short) 8); if ( cl_PMCTA_1stVisit_ANC!=null){ if ( cl_PMCTA_1stVisit_ANC.getCellType()==0) {PMCTA_1stVisit_ANC = "" + (int) cl_PMCTA_1stVisit_ANC.getNumericCellValue();} else if (cl_PMCTA_1stVisit_ANC.getCellType()==1) {PMCTA_1stVisit_ANC=cl_PMCTA_1stVisit_ANC.getStringCellValue();}  else if (cl_PMCTA_1stVisit_ANC.getCellType()== 2) {PMCTA_1stVisit_ANC=cl_PMCTA_1stVisit_ANC.getRawValue();  } else { PMCTA_1stVisit_ANC = "0";}}
XSSFCell cl_PMCTA_ReVisit_ANC = worksheet.getRow(8).getCell((short) 8); if ( cl_PMCTA_ReVisit_ANC!=null){ if ( cl_PMCTA_ReVisit_ANC.getCellType()==0) {PMCTA_ReVisit_ANC = "" + (int) cl_PMCTA_ReVisit_ANC.getNumericCellValue();} else if (cl_PMCTA_ReVisit_ANC.getCellType()==1) {PMCTA_ReVisit_ANC=cl_PMCTA_ReVisit_ANC.getStringCellValue();}  else if (cl_PMCTA_ReVisit_ANC.getCellType()== 2) {PMCTA_ReVisit_ANC=cl_PMCTA_ReVisit_ANC.getRawValue();  } else { PMCTA_ReVisit_ANC = "0";}}
XSSFCell cl_PMCTANCClientsT = worksheet.getRow(8).getCell((short) 9); if ( cl_PMCTANCClientsT!=null){ if ( cl_PMCTANCClientsT.getCellType()==0) {PMCTANCClientsT = "" + (int) cl_PMCTANCClientsT.getNumericCellValue();} else if (cl_PMCTANCClientsT.getCellType()==1) {PMCTANCClientsT=cl_PMCTANCClientsT.getStringCellValue();}  else if (cl_PMCTANCClientsT.getCellType()== 2) {PMCTANCClientsT=cl_PMCTANCClientsT.getRawValue();  } else { PMCTANCClientsT = "0";}}
XSSFCell cl_PMCTIPT1 = worksheet.getRow(9).getCell((short) 9); if ( cl_PMCTIPT1!=null){ if ( cl_PMCTIPT1.getCellType()==0) {PMCTIPT1 = "" + (int) cl_PMCTIPT1.getNumericCellValue();} else if (cl_PMCTIPT1.getCellType()==1) {PMCTIPT1=cl_PMCTIPT1.getStringCellValue();}  else if (cl_PMCTIPT1.getCellType()== 2) {PMCTIPT1=cl_PMCTIPT1.getRawValue();  } else { PMCTIPT1 = "0";}}
XSSFCell cl_PMCTIPT2 = worksheet.getRow(10).getCell((short) 9); if ( cl_PMCTIPT2!=null){ if ( cl_PMCTIPT2.getCellType()==0) {PMCTIPT2 = "" + (int) cl_PMCTIPT2.getNumericCellValue();} else if (cl_PMCTIPT2.getCellType()==1) {PMCTIPT2=cl_PMCTIPT2.getStringCellValue();}  else if (cl_PMCTIPT2.getCellType()== 2) {PMCTIPT2=cl_PMCTIPT2.getRawValue();  } else { PMCTIPT2 = "0";}}
XSSFCell cl_PMCTHB11 = worksheet.getRow(11).getCell((short) 9); if ( cl_PMCTHB11!=null){ if ( cl_PMCTHB11.getCellType()==0) {PMCTHB11 = "" + (int) cl_PMCTHB11.getNumericCellValue();} else if (cl_PMCTHB11.getCellType()==1) {PMCTHB11=cl_PMCTHB11.getStringCellValue();}  else if (cl_PMCTHB11.getCellType()== 2) {PMCTHB11=cl_PMCTHB11.getRawValue();  } else { PMCTHB11 = "0";}}
XSSFCell cl_PMCTANCClients4 = worksheet.getRow(12).getCell((short) 9); if ( cl_PMCTANCClients4!=null){ if ( cl_PMCTANCClients4.getCellType()==0) {PMCTANCClients4 = "" + (int) cl_PMCTANCClients4.getNumericCellValue();} else if (cl_PMCTANCClients4.getCellType()==1) {PMCTANCClients4=cl_PMCTANCClients4.getStringCellValue();}  else if (cl_PMCTANCClients4.getCellType()== 2) {PMCTANCClients4=cl_PMCTANCClients4.getRawValue();  } else { PMCTANCClients4 = "0";}}
XSSFCell cl_PMCTITN1 = worksheet.getRow(13).getCell((short) 9); if ( cl_PMCTITN1!=null){ if ( cl_PMCTITN1.getCellType()==0) {PMCTITN1 = "" + (int) cl_PMCTITN1.getNumericCellValue();} else if (cl_PMCTITN1.getCellType()==1) {PMCTITN1=cl_PMCTITN1.getStringCellValue();}  else if (cl_PMCTITN1.getCellType()== 2) {PMCTITN1=cl_PMCTITN1.getRawValue();  } else { PMCTITN1 = "0";}}
XSSFCell cl_PMCTITN = worksheet.getRow(14).getCell((short) 9); if ( cl_PMCTITN!=null){ if ( cl_PMCTITN.getCellType()==0) {PMCTITN = "" + (int) cl_PMCTITN.getNumericCellValue();} else if (cl_PMCTITN.getCellType()==1) {PMCTITN=cl_PMCTITN.getStringCellValue();}  else if (cl_PMCTITN.getCellType()== 2) {PMCTITN=cl_PMCTITN.getRawValue();  } else { PMCTITN = "0";}}
XSSFCell cl_PMTCTSYPHILISTES = worksheet.getRow(15).getCell((short) 9); if ( cl_PMTCTSYPHILISTES!=null){ if ( cl_PMTCTSYPHILISTES.getCellType()==0) {PMTCTSYPHILISTES = "" + (int) cl_PMTCTSYPHILISTES.getNumericCellValue();} else if (cl_PMTCTSYPHILISTES.getCellType()==1) {PMTCTSYPHILISTES=cl_PMTCTSYPHILISTES.getStringCellValue();}  else if (cl_PMTCTSYPHILISTES.getCellType()== 2) {PMTCTSYPHILISTES=cl_PMTCTSYPHILISTES.getRawValue();  } else { PMTCTSYPHILISTES = "0";}}
XSSFCell cl_PMTCTSYPHILISPOS = worksheet.getRow(16).getCell((short) 9); if ( cl_PMTCTSYPHILISPOS!=null){ if ( cl_PMTCTSYPHILISPOS.getCellType()==0) {PMTCTSYPHILISPOS = "" + (int) cl_PMTCTSYPHILISPOS.getNumericCellValue();} else if (cl_PMTCTSYPHILISPOS.getCellType()==1) {PMTCTSYPHILISPOS=cl_PMTCTSYPHILISPOS.getStringCellValue();}  else if (cl_PMTCTSYPHILISPOS.getCellType()== 2) {PMTCTSYPHILISPOS=cl_PMTCTSYPHILISPOS.getRawValue();  } else { PMTCTSYPHILISPOS = "0";}}
XSSFCell cl_PMTCTCOUNSELLEDFEED = worksheet.getRow(17).getCell((short) 9); if ( cl_PMTCTCOUNSELLEDFEED!=null){ if ( cl_PMTCTCOUNSELLEDFEED.getCellType()==0) {PMTCTCOUNSELLEDFEED = "" + (int) cl_PMTCTCOUNSELLEDFEED.getNumericCellValue();} else if (cl_PMTCTCOUNSELLEDFEED.getCellType()==1) {PMTCTCOUNSELLEDFEED=cl_PMTCTCOUNSELLEDFEED.getStringCellValue();}  else if (cl_PMTCTCOUNSELLEDFEED.getCellType()== 2) {PMTCTCOUNSELLEDFEED=cl_PMTCTCOUNSELLEDFEED.getRawValue();  } else { PMTCTCOUNSELLEDFEED = "0";}}
XSSFCell cl_PMTCTBREAST = worksheet.getRow(18).getCell((short) 9); if ( cl_PMTCTBREAST!=null){ if ( cl_PMTCTBREAST.getCellType()==0) {PMTCTBREAST = "" + (int) cl_PMTCTBREAST.getNumericCellValue();} else if (cl_PMTCTBREAST.getCellType()==1) {PMTCTBREAST=cl_PMTCTBREAST.getStringCellValue();}  else if (cl_PMTCTBREAST.getCellType()== 2) {PMTCTBREAST=cl_PMTCTBREAST.getRawValue();  } else { PMTCTBREAST = "0";}}
XSSFCell cl_PMTCTEXERCISE = worksheet.getRow(19).getCell((short) 9); if ( cl_PMTCTEXERCISE!=null){ if ( cl_PMTCTEXERCISE.getCellType()==0) {PMTCTEXERCISE = "" + (int) cl_PMTCTEXERCISE.getNumericCellValue();} else if (cl_PMTCTEXERCISE.getCellType()==1) {PMTCTEXERCISE=cl_PMTCTEXERCISE.getStringCellValue();}  else if (cl_PMTCTEXERCISE.getCellType()== 2) {PMTCTEXERCISE=cl_PMTCTEXERCISE.getRawValue();  } else { PMTCTEXERCISE = "0";}}
XSSFCell cl_PMTCTPREG10_14 = worksheet.getRow(20).getCell((short) 9); if ( cl_PMTCTPREG10_14!=null){ if ( cl_PMTCTPREG10_14.getCellType()==0) {PMTCTPREG10_14 = "" + (int) cl_PMTCTPREG10_14.getNumericCellValue();} else if (cl_PMTCTPREG10_14.getCellType()==1) {PMTCTPREG10_14=cl_PMTCTPREG10_14.getStringCellValue();}  else if (cl_PMTCTPREG10_14.getCellType()== 2) {PMTCTPREG10_14=cl_PMTCTPREG10_14.getRawValue();  } else { PMTCTPREG10_14 = "0";}}
XSSFCell cl_PMTCTPREG15_19 = worksheet.getRow(21).getCell((short) 9); if ( cl_PMTCTPREG15_19!=null){ if ( cl_PMTCTPREG15_19.getCellType()==0) {PMTCTPREG15_19 = "" + (int) cl_PMTCTPREG15_19.getNumericCellValue();} else if (cl_PMTCTPREG15_19.getCellType()==1) {PMTCTPREG15_19=cl_PMTCTPREG15_19.getStringCellValue();}  else if (cl_PMTCTPREG15_19.getCellType()== 2) {PMTCTPREG15_19=cl_PMTCTPREG15_19.getRawValue();  } else { PMTCTPREG15_19 = "0";}}
XSSFCell cl_PMTCTIRON = worksheet.getRow(22).getCell((short) 9); if ( cl_PMTCTIRON!=null){ if ( cl_PMTCTIRON.getCellType()==0) {PMTCTIRON = "" + (int) cl_PMTCTIRON.getNumericCellValue();} else if (cl_PMTCTIRON.getCellType()==1) {PMTCTIRON=cl_PMTCTIRON.getStringCellValue();}  else if (cl_PMTCTIRON.getCellType()== 2) {PMTCTIRON=cl_PMTCTIRON.getRawValue();  } else { PMTCTIRON = "0";}}
XSSFCell cl_PMTCTFOLIC = worksheet.getRow(23).getCell((short) 9); if ( cl_PMTCTFOLIC!=null){ if ( cl_PMTCTFOLIC.getCellType()==0) {PMTCTFOLIC = "" + (int) cl_PMTCTFOLIC.getNumericCellValue();} else if (cl_PMTCTFOLIC.getCellType()==1) {PMTCTFOLIC=cl_PMTCTFOLIC.getStringCellValue();}  else if (cl_PMTCTFOLIC.getCellType()== 2) {PMTCTFOLIC=cl_PMTCTFOLIC.getRawValue();  } else { PMTCTFOLIC = "0";}}
XSSFCell cl_PMTCTFERROUS = worksheet.getRow(24).getCell((short) 9); if ( cl_PMTCTFERROUS!=null){ if ( cl_PMTCTFERROUS.getCellType()==0) {PMTCTFERROUS = "" + (int) cl_PMTCTFERROUS.getNumericCellValue();} else if (cl_PMTCTFERROUS.getCellType()==1) {PMTCTFERROUS=cl_PMTCTFERROUS.getStringCellValue();}  else if (cl_PMTCTFERROUS.getCellType()== 2) {PMTCTFERROUS=cl_PMTCTFERROUS.getRawValue();  } else { PMTCTFERROUS = "0";}}
XSSFCell cl_MATNormalDelivery = worksheet.getRow(7).getCell((short) 16); if ( cl_MATNormalDelivery!=null){ if ( cl_MATNormalDelivery.getCellType()==0) {MATNormalDelivery = "" + (int) cl_MATNormalDelivery.getNumericCellValue();} else if (cl_MATNormalDelivery.getCellType()==1) {MATNormalDelivery=cl_MATNormalDelivery.getStringCellValue();}  else if (cl_MATNormalDelivery.getCellType()== 2) {MATNormalDelivery=cl_MATNormalDelivery.getRawValue();  } else { MATNormalDelivery = "0";}}
XSSFCell cl_MATCSection = worksheet.getRow(8).getCell((short) 16); if ( cl_MATCSection!=null){ if ( cl_MATCSection.getCellType()==0) {MATCSection = "" + (int) cl_MATCSection.getNumericCellValue();} else if (cl_MATCSection.getCellType()==1) {MATCSection=cl_MATCSection.getStringCellValue();}  else if (cl_MATCSection.getCellType()== 2) {MATCSection=cl_MATCSection.getRawValue();  } else { MATCSection = "0";}}
XSSFCell cl_MATBreech = worksheet.getRow(9).getCell((short) 16); if ( cl_MATBreech!=null){ if ( cl_MATBreech.getCellType()==0) {MATBreech = "" + (int) cl_MATBreech.getNumericCellValue();} else if (cl_MATBreech.getCellType()==1) {MATBreech=cl_MATBreech.getStringCellValue();}  else if (cl_MATBreech.getCellType()== 2) {MATBreech=cl_MATBreech.getRawValue();  } else { MATBreech = "0";}}
XSSFCell cl_MATAssistedVag = worksheet.getRow(10).getCell((short) 16); if ( cl_MATAssistedVag!=null){ if ( cl_MATAssistedVag.getCellType()==0) {MATAssistedVag = "" + (int) cl_MATAssistedVag.getNumericCellValue();} else if (cl_MATAssistedVag.getCellType()==1) {MATAssistedVag=cl_MATAssistedVag.getStringCellValue();}  else if (cl_MATAssistedVag.getCellType()== 2) {MATAssistedVag=cl_MATAssistedVag.getRawValue();  } else { MATAssistedVag = "0";}}
XSSFCell cl_MATDeliveryT = worksheet.getRow(11).getCell((short) 16); if ( cl_MATDeliveryT!=null){ if ( cl_MATDeliveryT.getCellType()==0) {MATDeliveryT = "" + (int) cl_MATDeliveryT.getNumericCellValue();} else if (cl_MATDeliveryT.getCellType()==1) {MATDeliveryT=cl_MATDeliveryT.getStringCellValue();}  else if (cl_MATDeliveryT.getCellType()== 2) {MATDeliveryT=cl_MATDeliveryT.getRawValue();  } else { MATDeliveryT = "0";}}
XSSFCell cl_MATLiveBirth = worksheet.getRow(12).getCell((short) 16); if ( cl_MATLiveBirth!=null){ if ( cl_MATLiveBirth.getCellType()==0) {MATLiveBirth = "" + (int) cl_MATLiveBirth.getNumericCellValue();} else if (cl_MATLiveBirth.getCellType()==1) {MATLiveBirth=cl_MATLiveBirth.getStringCellValue();}  else if (cl_MATLiveBirth.getCellType()== 2) {MATLiveBirth=cl_MATLiveBirth.getRawValue();  } else { MATLiveBirth = "0";}}
XSSFCell cl_MATFreshStillBirth = worksheet.getRow(13).getCell((short) 16); if ( cl_MATFreshStillBirth!=null){ if ( cl_MATFreshStillBirth.getCellType()==0) {MATFreshStillBirth = "" + (int) cl_MATFreshStillBirth.getNumericCellValue();} else if (cl_MATFreshStillBirth.getCellType()==1) {MATFreshStillBirth=cl_MATFreshStillBirth.getStringCellValue();}  else if (cl_MATFreshStillBirth.getCellType()== 2) {MATFreshStillBirth=cl_MATFreshStillBirth.getRawValue();  } else { MATFreshStillBirth = "0";}}
XSSFCell cl_MATMeceratedStillBirth = worksheet.getRow(14).getCell((short) 16); if ( cl_MATMeceratedStillBirth!=null){ if ( cl_MATMeceratedStillBirth.getCellType()==0) {MATMeceratedStillBirth = "" + (int) cl_MATMeceratedStillBirth.getNumericCellValue();} else if (cl_MATMeceratedStillBirth.getCellType()==1) {MATMeceratedStillBirth=cl_MATMeceratedStillBirth.getStringCellValue();}  else if (cl_MATMeceratedStillBirth.getCellType()== 2) {MATMeceratedStillBirth=cl_MATMeceratedStillBirth.getRawValue();  } else { MATMeceratedStillBirth = "0";}}
XSSFCell cl_MATDeformities = worksheet.getRow(15).getCell((short) 16); if ( cl_MATDeformities!=null){ if ( cl_MATDeformities.getCellType()==0) {MATDeformities = "" + (int) cl_MATDeformities.getNumericCellValue();} else if (cl_MATDeformities.getCellType()==1) {MATDeformities=cl_MATDeformities.getStringCellValue();}  else if (cl_MATDeformities.getCellType()== 2) {MATDeformities=cl_MATDeformities.getRawValue();  } else { MATDeformities = "0";}}
XSSFCell cl_MATLowAPGAR = worksheet.getRow(16).getCell((short) 16); if ( cl_MATLowAPGAR!=null){ if ( cl_MATLowAPGAR.getCellType()==0) {MATLowAPGAR = "" + (int) cl_MATLowAPGAR.getNumericCellValue();} else if (cl_MATLowAPGAR.getCellType()==1) {MATLowAPGAR=cl_MATLowAPGAR.getStringCellValue();}  else if (cl_MATLowAPGAR.getCellType()== 2) {MATLowAPGAR=cl_MATLowAPGAR.getRawValue();  } else { MATLowAPGAR = "0";}}
XSSFCell cl_MATWeight2500 = worksheet.getRow(17).getCell((short) 16); if ( cl_MATWeight2500!=null){ if ( cl_MATWeight2500.getCellType()==0) {MATWeight2500 = "" + (int) cl_MATWeight2500.getNumericCellValue();} else if (cl_MATWeight2500.getCellType()==1) {MATWeight2500=cl_MATWeight2500.getStringCellValue();}  else if (cl_MATWeight2500.getCellType()== 2) {MATWeight2500=cl_MATWeight2500.getRawValue();  } else { MATWeight2500 = "0";}}
XSSFCell cl_MATTetracycline = worksheet.getRow(18).getCell((short) 16); if ( cl_MATTetracycline!=null){ if ( cl_MATTetracycline.getCellType()==0) {MATTetracycline = "" + (int) cl_MATTetracycline.getNumericCellValue();} else if (cl_MATTetracycline.getCellType()==1) {MATTetracycline=cl_MATTetracycline.getStringCellValue();}  else if (cl_MATTetracycline.getCellType()== 2) {MATTetracycline=cl_MATTetracycline.getRawValue();  } else { MATTetracycline = "0";}}
XSSFCell cl_MATPreTerm = worksheet.getRow(19).getCell((short) 16); if ( cl_MATPreTerm!=null){ if ( cl_MATPreTerm.getCellType()==0) {MATPreTerm = "" + (int) cl_MATPreTerm.getNumericCellValue();} else if (cl_MATPreTerm.getCellType()==1) {MATPreTerm=cl_MATPreTerm.getStringCellValue();}  else if (cl_MATPreTerm.getCellType()== 2) {MATPreTerm=cl_MATPreTerm.getRawValue();  } else { MATPreTerm = "0";}}
XSSFCell cl_MATDischargealive = worksheet.getRow(20).getCell((short) 16); if ( cl_MATDischargealive!=null){ if ( cl_MATDischargealive.getCellType()==0) {MATDischargealive = "" + (int) cl_MATDischargealive.getNumericCellValue();} else if (cl_MATDischargealive.getCellType()==1) {MATDischargealive=cl_MATDischargealive.getStringCellValue();}  else if (cl_MATDischargealive.getCellType()== 2) {MATDischargealive=cl_MATDischargealive.getRawValue();  } else { MATDischargealive = "0";}}
XSSFCell cl_MATbreastfeeding1 = worksheet.getRow(21).getCell((short) 16); if ( cl_MATbreastfeeding1!=null){ if ( cl_MATbreastfeeding1.getCellType()==0) {MATbreastfeeding1 = "" + (int) cl_MATbreastfeeding1.getNumericCellValue();} else if (cl_MATbreastfeeding1.getCellType()==1) {MATbreastfeeding1=cl_MATbreastfeeding1.getStringCellValue();}  else if (cl_MATbreastfeeding1.getCellType()== 2) {MATbreastfeeding1=cl_MATbreastfeeding1.getRawValue();  } else { MATbreastfeeding1 = "0";}}
XSSFCell cl_MATDeliveriesPos = worksheet.getRow(22).getCell((short) 16); if ( cl_MATDeliveriesPos!=null){ if ( cl_MATDeliveriesPos.getCellType()==0) {MATDeliveriesPos = "" + (int) cl_MATDeliveriesPos.getNumericCellValue();} else if (cl_MATDeliveriesPos.getCellType()==1) {MATDeliveriesPos=cl_MATDeliveriesPos.getStringCellValue();}  else if (cl_MATDeliveriesPos.getCellType()== 2) {MATDeliveriesPos=cl_MATDeliveriesPos.getRawValue();  } else { MATDeliveriesPos = "0";}}
XSSFCell cl_MATNeoNatalD = worksheet.getRow(23).getCell((short) 16); if ( cl_MATNeoNatalD!=null){ if ( cl_MATNeoNatalD.getCellType()==0) {MATNeoNatalD = "" + (int) cl_MATNeoNatalD.getNumericCellValue();} else if (cl_MATNeoNatalD.getCellType()==1) {MATNeoNatalD=cl_MATNeoNatalD.getStringCellValue();}  else if (cl_MATNeoNatalD.getCellType()== 2) {MATNeoNatalD=cl_MATNeoNatalD.getRawValue();  } else { MATNeoNatalD = "0";}}
XSSFCell cl_MATMaternalD10_19 = worksheet.getRow(24).getCell((short) 16); if ( cl_MATMaternalD10_19!=null){ if ( cl_MATMaternalD10_19.getCellType()==0) {MATMaternalD10_19 = "" + (int) cl_MATMaternalD10_19.getNumericCellValue();} else if (cl_MATMaternalD10_19.getCellType()==1) {MATMaternalD10_19=cl_MATMaternalD10_19.getStringCellValue();}  else if (cl_MATMaternalD10_19.getCellType()== 2) {MATMaternalD10_19=cl_MATMaternalD10_19.getRawValue();  } else { MATMaternalD10_19 = "0";}}
XSSFCell cl_MATMaternalD = worksheet.getRow(25).getCell((short) 16); if ( cl_MATMaternalD!=null){ if ( cl_MATMaternalD.getCellType()==0) {MATMaternalD = "" + (int) cl_MATMaternalD.getNumericCellValue();} else if (cl_MATMaternalD.getCellType()==1) {MATMaternalD=cl_MATMaternalD.getStringCellValue();}  else if (cl_MATMaternalD.getCellType()== 2) {MATMaternalD=cl_MATMaternalD.getRawValue();  } else { MATMaternalD = "0";}}
XSSFCell cl_MATMaternalDAudited = worksheet.getRow(26).getCell((short) 16); if ( cl_MATMaternalDAudited!=null){ if ( cl_MATMaternalDAudited.getCellType()==0) {MATMaternalDAudited = "" + (int) cl_MATMaternalDAudited.getNumericCellValue();} else if (cl_MATMaternalDAudited.getCellType()==1) {MATMaternalDAudited=cl_MATMaternalDAudited.getStringCellValue();}  else if (cl_MATMaternalDAudited.getCellType()== 2) {MATMaternalDAudited=cl_MATMaternalDAudited.getRawValue();  } else { MATMaternalDAudited = "0";}}
XSSFCell cl_MATAPHAlive = worksheet.getRow(28).getCell((short) 15); if ( cl_MATAPHAlive!=null){ if ( cl_MATAPHAlive.getCellType()==0) {MATAPHAlive = "" + (int) cl_MATAPHAlive.getNumericCellValue();} else if (cl_MATAPHAlive.getCellType()==1) {MATAPHAlive=cl_MATAPHAlive.getStringCellValue();}  else if (cl_MATAPHAlive.getCellType()== 2) {MATAPHAlive=cl_MATAPHAlive.getRawValue();  } else { MATAPHAlive = "0";}}
XSSFCell cl_MATAPHDead = worksheet.getRow(28).getCell((short) 16); if ( cl_MATAPHDead!=null){ if ( cl_MATAPHDead.getCellType()==0) {MATAPHDead = "" + (int) cl_MATAPHDead.getNumericCellValue();} else if (cl_MATAPHDead.getCellType()==1) {MATAPHDead=cl_MATAPHDead.getStringCellValue();}  else if (cl_MATAPHDead.getCellType()== 2) {MATAPHDead=cl_MATAPHDead.getRawValue();  } else { MATAPHDead = "0";}}
XSSFCell cl_MATPPHAlive = worksheet.getRow(29).getCell((short) 15); if ( cl_MATPPHAlive!=null){ if ( cl_MATPPHAlive.getCellType()==0) {MATPPHAlive = "" + (int) cl_MATPPHAlive.getNumericCellValue();} else if (cl_MATPPHAlive.getCellType()==1) {MATPPHAlive=cl_MATPPHAlive.getStringCellValue();}  else if (cl_MATPPHAlive.getCellType()== 2) {MATPPHAlive=cl_MATPPHAlive.getRawValue();  } else { MATPPHAlive = "0";}}
XSSFCell cl_MATPPHDead = worksheet.getRow(29).getCell((short) 16); if ( cl_MATPPHDead!=null){ if ( cl_MATPPHDead.getCellType()==0) {MATPPHDead = "" + (int) cl_MATPPHDead.getNumericCellValue();} else if (cl_MATPPHDead.getCellType()==1) {MATPPHDead=cl_MATPPHDead.getStringCellValue();}  else if (cl_MATPPHDead.getCellType()== 2) {MATPPHDead=cl_MATPPHDead.getRawValue();  } else { MATPPHDead = "0";}}
XSSFCell cl_MATEclampAlive = worksheet.getRow(30).getCell((short) 15); if ( cl_MATEclampAlive!=null){ if ( cl_MATEclampAlive.getCellType()==0) {MATEclampAlive = "" + (int) cl_MATEclampAlive.getNumericCellValue();} else if (cl_MATEclampAlive.getCellType()==1) {MATEclampAlive=cl_MATEclampAlive.getStringCellValue();}  else if (cl_MATEclampAlive.getCellType()== 2) {MATEclampAlive=cl_MATEclampAlive.getRawValue();  } else { MATEclampAlive = "0";}}
XSSFCell cl_MATEclampDead = worksheet.getRow(30).getCell((short) 16); if ( cl_MATEclampDead!=null){ if ( cl_MATEclampDead.getCellType()==0) {MATEclampDead = "" + (int) cl_MATEclampDead.getNumericCellValue();} else if (cl_MATEclampDead.getCellType()==1) {MATEclampDead=cl_MATEclampDead.getStringCellValue();}  else if (cl_MATEclampDead.getCellType()== 2) {MATEclampDead=cl_MATEclampDead.getRawValue();  } else { MATEclampDead = "0";}}
XSSFCell cl_MATRupUtAlive = worksheet.getRow(31).getCell((short) 15); if ( cl_MATRupUtAlive!=null){ if ( cl_MATRupUtAlive.getCellType()==0) {MATRupUtAlive = "" + (int) cl_MATRupUtAlive.getNumericCellValue();} else if (cl_MATRupUtAlive.getCellType()==1) {MATRupUtAlive=cl_MATRupUtAlive.getStringCellValue();}  else if (cl_MATRupUtAlive.getCellType()== 2) {MATRupUtAlive=cl_MATRupUtAlive.getRawValue();  } else { MATRupUtAlive = "0";}}
XSSFCell cl_MATRupUtDead = worksheet.getRow(31).getCell((short) 16); if ( cl_MATRupUtDead!=null){ if ( cl_MATRupUtDead.getCellType()==0) {MATRupUtDead = "" + (int) cl_MATRupUtDead.getNumericCellValue();} else if (cl_MATRupUtDead.getCellType()==1) {MATRupUtDead=cl_MATRupUtDead.getStringCellValue();}  else if (cl_MATRupUtDead.getCellType()== 2) {MATRupUtDead=cl_MATRupUtDead.getRawValue();  } else { MATRupUtDead = "0";}}
XSSFCell cl_MATObstrLaborAlive = worksheet.getRow(32).getCell((short) 15); if ( cl_MATObstrLaborAlive!=null){ if ( cl_MATObstrLaborAlive.getCellType()==0) {MATObstrLaborAlive = "" + (int) cl_MATObstrLaborAlive.getNumericCellValue();} else if (cl_MATObstrLaborAlive.getCellType()==1) {MATObstrLaborAlive=cl_MATObstrLaborAlive.getStringCellValue();}  else if (cl_MATObstrLaborAlive.getCellType()== 2) {MATObstrLaborAlive=cl_MATObstrLaborAlive.getRawValue();  } else { MATObstrLaborAlive = "0";}}
XSSFCell cl_MATObstrLaborDead = worksheet.getRow(32).getCell((short) 16); if ( cl_MATObstrLaborDead!=null){ if ( cl_MATObstrLaborDead.getCellType()==0) {MATObstrLaborDead = "" + (int) cl_MATObstrLaborDead.getNumericCellValue();} else if (cl_MATObstrLaborDead.getCellType()==1) {MATObstrLaborDead=cl_MATObstrLaborDead.getStringCellValue();}  else if (cl_MATObstrLaborDead.getCellType()== 2) {MATObstrLaborDead=cl_MATObstrLaborDead.getRawValue();  } else { MATObstrLaborDead = "0";}}
XSSFCell cl_MATSepsisAlive = worksheet.getRow(33).getCell((short) 15); if ( cl_MATSepsisAlive!=null){ if ( cl_MATSepsisAlive.getCellType()==0) {MATSepsisAlive = "" + (int) cl_MATSepsisAlive.getNumericCellValue();} else if (cl_MATSepsisAlive.getCellType()==1) {MATSepsisAlive=cl_MATSepsisAlive.getStringCellValue();}  else if (cl_MATSepsisAlive.getCellType()== 2) {MATSepsisAlive=cl_MATSepsisAlive.getRawValue();  } else { MATSepsisAlive = "0";}}
XSSFCell cl_MATSepsisDead = worksheet.getRow(33).getCell((short) 16); if ( cl_MATSepsisDead!=null){ if ( cl_MATSepsisDead.getCellType()==0) {MATSepsisDead = "" + (int) cl_MATSepsisDead.getNumericCellValue();} else if (cl_MATSepsisDead.getCellType()==1) {MATSepsisDead=cl_MATSepsisDead.getStringCellValue();}  else if (cl_MATSepsisDead.getCellType()== 2) {MATSepsisDead=cl_MATSepsisDead.getRawValue();  } else { MATSepsisDead = "0";}}
XSSFCell cl_MATREFFromOtherFacility = worksheet.getRow(35).getCell((short) 15); if ( cl_MATREFFromOtherFacility!=null){ if ( cl_MATREFFromOtherFacility.getCellType()==0) {MATREFFromOtherFacility = "" + (int) cl_MATREFFromOtherFacility.getNumericCellValue();} else if (cl_MATREFFromOtherFacility.getCellType()==1) {MATREFFromOtherFacility=cl_MATREFFromOtherFacility.getStringCellValue();}  else if (cl_MATREFFromOtherFacility.getCellType()== 2) {MATREFFromOtherFacility=cl_MATREFFromOtherFacility.getRawValue();  } else { MATREFFromOtherFacility = "0";}}
XSSFCell cl_MATREFFromCU = worksheet.getRow(36).getCell((short) 15); if ( cl_MATREFFromCU!=null){ if ( cl_MATREFFromCU.getCellType()==0) {MATREFFromCU = "" + (int) cl_MATREFFromCU.getNumericCellValue();} else if (cl_MATREFFromCU.getCellType()==1) {MATREFFromCU=cl_MATREFFromCU.getStringCellValue();}  else if (cl_MATREFFromCU.getCellType()== 2) {MATREFFromCU=cl_MATREFFromCU.getRawValue();  } else { MATREFFromCU = "0";}}
XSSFCell cl_MATREFToOtherFacility = worksheet.getRow(37).getCell((short) 15); if ( cl_MATREFToOtherFacility!=null){ if ( cl_MATREFToOtherFacility.getCellType()==0) {MATREFToOtherFacility = "" + (int) cl_MATREFToOtherFacility.getNumericCellValue();} else if (cl_MATREFToOtherFacility.getCellType()==1) {MATREFToOtherFacility=cl_MATREFToOtherFacility.getStringCellValue();}  else if (cl_MATREFToOtherFacility.getCellType()== 2) {MATREFToOtherFacility=cl_MATREFToOtherFacility.getRawValue();  } else { MATREFToOtherFacility = "0";}}
XSSFCell cl_MATREFToCU = worksheet.getRow(38).getCell((short) 15); if ( cl_MATREFToCU!=null){ if ( cl_MATREFToCU.getCellType()==0) {MATREFToCU = "" + (int) cl_MATREFToCU.getNumericCellValue();} else if (cl_MATREFToCU.getCellType()==1) {MATREFToCU=cl_MATREFToCU.getStringCellValue();}  else if (cl_MATREFToCU.getCellType()== 2) {MATREFToCU=cl_MATREFToCU.getRawValue();  } else { MATREFToCU = "0";}}
XSSFCell cl_SGBVRape72_0_9 = worksheet.getRow(27).getCell((short) 6); if ( cl_SGBVRape72_0_9!=null){ if ( cl_SGBVRape72_0_9.getCellType()==0) {SGBVRape72_0_9 = "" + (int) cl_SGBVRape72_0_9.getNumericCellValue();} else if (cl_SGBVRape72_0_9.getCellType()==1) {SGBVRape72_0_9=cl_SGBVRape72_0_9.getStringCellValue();}  else if (cl_SGBVRape72_0_9.getCellType()== 2) {SGBVRape72_0_9=cl_SGBVRape72_0_9.getRawValue();  } else { SGBVRape72_0_9 = "0";}}
XSSFCell cl_SGBVRape72_10_17 = worksheet.getRow(27).getCell((short) 7); if ( cl_SGBVRape72_10_17!=null){ if ( cl_SGBVRape72_10_17.getCellType()==0) {SGBVRape72_10_17 = "" + (int) cl_SGBVRape72_10_17.getNumericCellValue();} else if (cl_SGBVRape72_10_17.getCellType()==1) {SGBVRape72_10_17=cl_SGBVRape72_10_17.getStringCellValue();}  else if (cl_SGBVRape72_10_17.getCellType()== 2) {SGBVRape72_10_17=cl_SGBVRape72_10_17.getRawValue();  } else { SGBVRape72_10_17 = "0";}}
XSSFCell cl_SGBVRape72_18_49 = worksheet.getRow(27).getCell((short) 8); if ( cl_SGBVRape72_18_49!=null){ if ( cl_SGBVRape72_18_49.getCellType()==0) {SGBVRape72_18_49 = "" + (int) cl_SGBVRape72_18_49.getNumericCellValue();} else if (cl_SGBVRape72_18_49.getCellType()==1) {SGBVRape72_18_49=cl_SGBVRape72_18_49.getStringCellValue();}  else if (cl_SGBVRape72_18_49.getCellType()== 2) {SGBVRape72_18_49=cl_SGBVRape72_18_49.getRawValue();  } else { SGBVRape72_18_49 = "0";}}
XSSFCell cl_SGBVRape72_50 = worksheet.getRow(27).getCell((short) 9); if ( cl_SGBVRape72_50!=null){ if ( cl_SGBVRape72_50.getCellType()==0) {SGBVRape72_50 = "" + (int) cl_SGBVRape72_50.getNumericCellValue();} else if (cl_SGBVRape72_50.getCellType()==1) {SGBVRape72_50=cl_SGBVRape72_50.getStringCellValue();}  else if (cl_SGBVRape72_50.getCellType()== 2) {SGBVRape72_50=cl_SGBVRape72_50.getRawValue();  } else { SGBVRape72_50 = "0";}}
XSSFCell cl_SGBVinitPEP0_9 = worksheet.getRow(28).getCell((short) 6); if ( cl_SGBVinitPEP0_9!=null){ if ( cl_SGBVinitPEP0_9.getCellType()==0) {SGBVinitPEP0_9 = "" + (int) cl_SGBVinitPEP0_9.getNumericCellValue();} else if (cl_SGBVinitPEP0_9.getCellType()==1) {SGBVinitPEP0_9=cl_SGBVinitPEP0_9.getStringCellValue();}  else if (cl_SGBVinitPEP0_9.getCellType()== 2) {SGBVinitPEP0_9=cl_SGBVinitPEP0_9.getRawValue();  } else { SGBVinitPEP0_9 = "0";}}
XSSFCell cl_SGBVinitPEP10_17 = worksheet.getRow(28).getCell((short) 7); if ( cl_SGBVinitPEP10_17!=null){ if ( cl_SGBVinitPEP10_17.getCellType()==0) {SGBVinitPEP10_17 = "" + (int) cl_SGBVinitPEP10_17.getNumericCellValue();} else if (cl_SGBVinitPEP10_17.getCellType()==1) {SGBVinitPEP10_17=cl_SGBVinitPEP10_17.getStringCellValue();}  else if (cl_SGBVinitPEP10_17.getCellType()== 2) {SGBVinitPEP10_17=cl_SGBVinitPEP10_17.getRawValue();  } else { SGBVinitPEP10_17 = "0";}}
XSSFCell cl_SGBVinitPEP18_49 = worksheet.getRow(28).getCell((short) 8); if ( cl_SGBVinitPEP18_49!=null){ if ( cl_SGBVinitPEP18_49.getCellType()==0) {SGBVinitPEP18_49 = "" + (int) cl_SGBVinitPEP18_49.getNumericCellValue();} else if (cl_SGBVinitPEP18_49.getCellType()==1) {SGBVinitPEP18_49=cl_SGBVinitPEP18_49.getStringCellValue();}  else if (cl_SGBVinitPEP18_49.getCellType()== 2) {SGBVinitPEP18_49=cl_SGBVinitPEP18_49.getRawValue();  } else { SGBVinitPEP18_49 = "0";}}
XSSFCell cl_SGBVinitPEP50 = worksheet.getRow(28).getCell((short) 9); if ( cl_SGBVinitPEP50!=null){ if ( cl_SGBVinitPEP50.getCellType()==0) {SGBVinitPEP50 = "" + (int) cl_SGBVinitPEP50.getNumericCellValue();} else if (cl_SGBVinitPEP50.getCellType()==1) {SGBVinitPEP50=cl_SGBVinitPEP50.getStringCellValue();}  else if (cl_SGBVinitPEP50.getCellType()== 2) {SGBVinitPEP50=cl_SGBVinitPEP50.getRawValue();  } else { SGBVinitPEP50 = "0";}}
XSSFCell cl_SGBVcompPEP0_9 = worksheet.getRow(29).getCell((short) 6); if ( cl_SGBVcompPEP0_9!=null){ if ( cl_SGBVcompPEP0_9.getCellType()==0) {SGBVcompPEP0_9 = "" + (int) cl_SGBVcompPEP0_9.getNumericCellValue();} else if (cl_SGBVcompPEP0_9.getCellType()==1) {SGBVcompPEP0_9=cl_SGBVcompPEP0_9.getStringCellValue();}  else if (cl_SGBVcompPEP0_9.getCellType()== 2) {SGBVcompPEP0_9=cl_SGBVcompPEP0_9.getRawValue();  } else { SGBVcompPEP0_9 = "0";}}
XSSFCell cl_SGBVcompPEP10_17 = worksheet.getRow(29).getCell((short) 7); if ( cl_SGBVcompPEP10_17!=null){ if ( cl_SGBVcompPEP10_17.getCellType()==0) {SGBVcompPEP10_17 = "" + (int) cl_SGBVcompPEP10_17.getNumericCellValue();} else if (cl_SGBVcompPEP10_17.getCellType()==1) {SGBVcompPEP10_17=cl_SGBVcompPEP10_17.getStringCellValue();}  else if (cl_SGBVcompPEP10_17.getCellType()== 2) {SGBVcompPEP10_17=cl_SGBVcompPEP10_17.getRawValue();  } else { SGBVcompPEP10_17 = "0";}}
XSSFCell cl_SGBVcompPEP18_49 = worksheet.getRow(29).getCell((short) 8); if ( cl_SGBVcompPEP18_49!=null){ if ( cl_SGBVcompPEP18_49.getCellType()==0) {SGBVcompPEP18_49 = "" + (int) cl_SGBVcompPEP18_49.getNumericCellValue();} else if (cl_SGBVcompPEP18_49.getCellType()==1) {SGBVcompPEP18_49=cl_SGBVcompPEP18_49.getStringCellValue();}  else if (cl_SGBVcompPEP18_49.getCellType()== 2) {SGBVcompPEP18_49=cl_SGBVcompPEP18_49.getRawValue();  } else { SGBVcompPEP18_49 = "0";}}
XSSFCell cl_SGBVcompPEP50 = worksheet.getRow(29).getCell((short) 9); if ( cl_SGBVcompPEP50!=null){ if ( cl_SGBVcompPEP50.getCellType()==0) {SGBVcompPEP50 = "" + (int) cl_SGBVcompPEP50.getNumericCellValue();} else if (cl_SGBVcompPEP50.getCellType()==1) {SGBVcompPEP50=cl_SGBVcompPEP50.getStringCellValue();}  else if (cl_SGBVcompPEP50.getCellType()== 2) {SGBVcompPEP50=cl_SGBVcompPEP50.getRawValue();  } else { SGBVcompPEP50 = "0";}}
XSSFCell cl_SGBVPregnant0_9 = worksheet.getRow(30).getCell((short) 6); if ( cl_SGBVPregnant0_9!=null){ if ( cl_SGBVPregnant0_9.getCellType()==0) {SGBVPregnant0_9 = "" + (int) cl_SGBVPregnant0_9.getNumericCellValue();} else if (cl_SGBVPregnant0_9.getCellType()==1) {SGBVPregnant0_9=cl_SGBVPregnant0_9.getStringCellValue();}  else if (cl_SGBVPregnant0_9.getCellType()== 2) {SGBVPregnant0_9=cl_SGBVPregnant0_9.getRawValue();  } else { SGBVPregnant0_9 = "0";}}
XSSFCell cl_SGBVPregnant10_17 = worksheet.getRow(30).getCell((short) 7); if ( cl_SGBVPregnant10_17!=null){ if ( cl_SGBVPregnant10_17.getCellType()==0) {SGBVPregnant10_17 = "" + (int) cl_SGBVPregnant10_17.getNumericCellValue();} else if (cl_SGBVPregnant10_17.getCellType()==1) {SGBVPregnant10_17=cl_SGBVPregnant10_17.getStringCellValue();}  else if (cl_SGBVPregnant10_17.getCellType()== 2) {SGBVPregnant10_17=cl_SGBVPregnant10_17.getRawValue();  } else { SGBVPregnant10_17 = "0";}}
XSSFCell cl_SGBVPregnant18_49 = worksheet.getRow(30).getCell((short) 8); if ( cl_SGBVPregnant18_49!=null){ if ( cl_SGBVPregnant18_49.getCellType()==0) {SGBVPregnant18_49 = "" + (int) cl_SGBVPregnant18_49.getNumericCellValue();} else if (cl_SGBVPregnant18_49.getCellType()==1) {SGBVPregnant18_49=cl_SGBVPregnant18_49.getStringCellValue();}  else if (cl_SGBVPregnant18_49.getCellType()== 2) {SGBVPregnant18_49=cl_SGBVPregnant18_49.getRawValue();  } else { SGBVPregnant18_49 = "0";}}
XSSFCell cl_SGBVPregnant50 = worksheet.getRow(30).getCell((short) 9); if ( cl_SGBVPregnant50!=null){ if ( cl_SGBVPregnant50.getCellType()==0) {SGBVPregnant50 = "" + (int) cl_SGBVPregnant50.getNumericCellValue();} else if (cl_SGBVPregnant50.getCellType()==1) {SGBVPregnant50=cl_SGBVPregnant50.getStringCellValue();}  else if (cl_SGBVPregnant50.getCellType()== 2) {SGBVPregnant50=cl_SGBVPregnant50.getRawValue();  } else { SGBVPregnant50 = "0";}}
XSSFCell cl_SGBVseroconverting0_9 = worksheet.getRow(31).getCell((short) 6); if ( cl_SGBVseroconverting0_9!=null){ if ( cl_SGBVseroconverting0_9.getCellType()==0) {SGBVseroconverting0_9 = "" + (int) cl_SGBVseroconverting0_9.getNumericCellValue();} else if (cl_SGBVseroconverting0_9.getCellType()==1) {SGBVseroconverting0_9=cl_SGBVseroconverting0_9.getStringCellValue();}  else if (cl_SGBVseroconverting0_9.getCellType()== 2) {SGBVseroconverting0_9=cl_SGBVseroconverting0_9.getRawValue();  } else { SGBVseroconverting0_9 = "0";}}
XSSFCell cl_SGBVseroconverting10_17 = worksheet.getRow(31).getCell((short) 7); if ( cl_SGBVseroconverting10_17!=null){ if ( cl_SGBVseroconverting10_17.getCellType()==0) {SGBVseroconverting10_17 = "" + (int) cl_SGBVseroconverting10_17.getNumericCellValue();} else if (cl_SGBVseroconverting10_17.getCellType()==1) {SGBVseroconverting10_17=cl_SGBVseroconverting10_17.getStringCellValue();}  else if (cl_SGBVseroconverting10_17.getCellType()== 2) {SGBVseroconverting10_17=cl_SGBVseroconverting10_17.getRawValue();  } else { SGBVseroconverting10_17 = "0";}}
XSSFCell cl_SGBVseroconverting18_49 = worksheet.getRow(31).getCell((short) 8); if ( cl_SGBVseroconverting18_49!=null){ if ( cl_SGBVseroconverting18_49.getCellType()==0) {SGBVseroconverting18_49 = "" + (int) cl_SGBVseroconverting18_49.getNumericCellValue();} else if (cl_SGBVseroconverting18_49.getCellType()==1) {SGBVseroconverting18_49=cl_SGBVseroconverting18_49.getStringCellValue();}  else if (cl_SGBVseroconverting18_49.getCellType()== 2) {SGBVseroconverting18_49=cl_SGBVseroconverting18_49.getRawValue();  } else { SGBVseroconverting18_49 = "0";}}
XSSFCell cl_SGBVseroconverting50 = worksheet.getRow(31).getCell((short) 9); if ( cl_SGBVseroconverting50!=null){ if ( cl_SGBVseroconverting50.getCellType()==0) {SGBVseroconverting50 = "" + (int) cl_SGBVseroconverting50.getNumericCellValue();} else if (cl_SGBVseroconverting50.getCellType()==1) {SGBVseroconverting50=cl_SGBVseroconverting50.getStringCellValue();}  else if (cl_SGBVseroconverting50.getCellType()== 2) {SGBVseroconverting50=cl_SGBVseroconverting50.getRawValue();  } else { SGBVseroconverting50 = "0";}}
XSSFCell cl_SGBVsurvivors0_9 = worksheet.getRow(32).getCell((short) 6); if ( cl_SGBVsurvivors0_9!=null){ if ( cl_SGBVsurvivors0_9.getCellType()==0) {SGBVsurvivors0_9 = "" + (int) cl_SGBVsurvivors0_9.getNumericCellValue();} else if (cl_SGBVsurvivors0_9.getCellType()==1) {SGBVsurvivors0_9=cl_SGBVsurvivors0_9.getStringCellValue();}  else if (cl_SGBVsurvivors0_9.getCellType()== 2) {SGBVsurvivors0_9=cl_SGBVsurvivors0_9.getRawValue();  } else { SGBVsurvivors0_9 = "0";}}
XSSFCell cl_SGBVsurvivors10_17 = worksheet.getRow(32).getCell((short) 7); if ( cl_SGBVsurvivors10_17!=null){ if ( cl_SGBVsurvivors10_17.getCellType()==0) {SGBVsurvivors10_17 = "" + (int) cl_SGBVsurvivors10_17.getNumericCellValue();} else if (cl_SGBVsurvivors10_17.getCellType()==1) {SGBVsurvivors10_17=cl_SGBVsurvivors10_17.getStringCellValue();}  else if (cl_SGBVsurvivors10_17.getCellType()== 2) {SGBVsurvivors10_17=cl_SGBVsurvivors10_17.getRawValue();  } else { SGBVsurvivors10_17 = "0";}}
XSSFCell cl_SGBVsurvivors18_49 = worksheet.getRow(32).getCell((short) 8); if ( cl_SGBVsurvivors18_49!=null){ if ( cl_SGBVsurvivors18_49.getCellType()==0) {SGBVsurvivors18_49 = "" + (int) cl_SGBVsurvivors18_49.getNumericCellValue();} else if (cl_SGBVsurvivors18_49.getCellType()==1) {SGBVsurvivors18_49=cl_SGBVsurvivors18_49.getStringCellValue();}  else if (cl_SGBVsurvivors18_49.getCellType()== 2) {SGBVsurvivors18_49=cl_SGBVsurvivors18_49.getRawValue();  } else { SGBVsurvivors18_49 = "0";}}
XSSFCell cl_SGBVsurvivors50 = worksheet.getRow(32).getCell((short) 9); if ( cl_SGBVsurvivors50!=null){ if ( cl_SGBVsurvivors50.getCellType()==0) {SGBVsurvivors50 = "" + (int) cl_SGBVsurvivors50.getNumericCellValue();} else if (cl_SGBVsurvivors50.getCellType()==1) {SGBVsurvivors50=cl_SGBVsurvivors50.getStringCellValue();}  else if (cl_SGBVsurvivors50.getCellType()== 2) {SGBVsurvivors50=cl_SGBVsurvivors50.getRawValue();  } else { SGBVsurvivors50 = "0";}}
XSSFCell cl_PAC10_19 = worksheet.getRow(41).getCell((short) 15); if ( cl_PAC10_19!=null){ if ( cl_PAC10_19.getCellType()==0) {PAC10_19 = "" + (int) cl_PAC10_19.getNumericCellValue();} else if (cl_PAC10_19.getCellType()==1) {PAC10_19=cl_PAC10_19.getStringCellValue();}  else if (cl_PAC10_19.getCellType()== 2) {PAC10_19=cl_PAC10_19.getRawValue();  } else { PAC10_19 = "0";}}
XSSFCell cl_PACT = worksheet.getRow(42).getCell((short) 15); if ( cl_PACT!=null){ if ( cl_PACT.getCellType()==0) {PACT = "" + (int) cl_PACT.getNumericCellValue();} else if (cl_PACT.getCellType()==1) {PACT=cl_PACT.getStringCellValue();}  else if (cl_PACT.getCellType()== 2) {PACT=cl_PACT.getRawValue();  } else { PACT = "0";}}
XSSFCell cl_CHANIS0_5NormalweightF = worksheet.getRow(46).getCell((short) 14); if ( cl_CHANIS0_5NormalweightF!=null){ if ( cl_CHANIS0_5NormalweightF.getCellType()==0) {CHANIS0_5NormalweightF = "" + (int) cl_CHANIS0_5NormalweightF.getNumericCellValue();} else if (cl_CHANIS0_5NormalweightF.getCellType()==1) {CHANIS0_5NormalweightF=cl_CHANIS0_5NormalweightF.getStringCellValue();}  else if (cl_CHANIS0_5NormalweightF.getCellType()== 2) {CHANIS0_5NormalweightF=cl_CHANIS0_5NormalweightF.getRawValue();  } else { CHANIS0_5NormalweightF = "0";}}
XSSFCell cl_CHANIS0_5NormalweightM = worksheet.getRow(46).getCell((short) 15); if ( cl_CHANIS0_5NormalweightM!=null){ if ( cl_CHANIS0_5NormalweightM.getCellType()==0) {CHANIS0_5NormalweightM = "" + (int) cl_CHANIS0_5NormalweightM.getNumericCellValue();} else if (cl_CHANIS0_5NormalweightM.getCellType()==1) {CHANIS0_5NormalweightM=cl_CHANIS0_5NormalweightM.getStringCellValue();}  else if (cl_CHANIS0_5NormalweightM.getCellType()== 2) {CHANIS0_5NormalweightM=cl_CHANIS0_5NormalweightM.getRawValue();  } else { CHANIS0_5NormalweightM = "0";}}
XSSFCell cl_CHANIS0_5NormalweightT = worksheet.getRow(46).getCell((short) 16); if ( cl_CHANIS0_5NormalweightT!=null){ if ( cl_CHANIS0_5NormalweightT.getCellType()==0) {CHANIS0_5NormalweightT = "" + (int) cl_CHANIS0_5NormalweightT.getNumericCellValue();} else if (cl_CHANIS0_5NormalweightT.getCellType()==1) {CHANIS0_5NormalweightT=cl_CHANIS0_5NormalweightT.getStringCellValue();}  else if (cl_CHANIS0_5NormalweightT.getCellType()== 2) {CHANIS0_5NormalweightT=cl_CHANIS0_5NormalweightT.getRawValue();  } else { CHANIS0_5NormalweightT = "0";}}
XSSFCell cl_CHANIS0_5UnderweightF = worksheet.getRow(47).getCell((short) 14); if ( cl_CHANIS0_5UnderweightF!=null){ if ( cl_CHANIS0_5UnderweightF.getCellType()==0) {CHANIS0_5UnderweightF = "" + (int) cl_CHANIS0_5UnderweightF.getNumericCellValue();} else if (cl_CHANIS0_5UnderweightF.getCellType()==1) {CHANIS0_5UnderweightF=cl_CHANIS0_5UnderweightF.getStringCellValue();}  else if (cl_CHANIS0_5UnderweightF.getCellType()== 2) {CHANIS0_5UnderweightF=cl_CHANIS0_5UnderweightF.getRawValue();  } else { CHANIS0_5UnderweightF = "0";}}
XSSFCell cl_CHANIS0_5UnderweightM = worksheet.getRow(47).getCell((short) 15); if ( cl_CHANIS0_5UnderweightM!=null){ if ( cl_CHANIS0_5UnderweightM.getCellType()==0) {CHANIS0_5UnderweightM = "" + (int) cl_CHANIS0_5UnderweightM.getNumericCellValue();} else if (cl_CHANIS0_5UnderweightM.getCellType()==1) {CHANIS0_5UnderweightM=cl_CHANIS0_5UnderweightM.getStringCellValue();}  else if (cl_CHANIS0_5UnderweightM.getCellType()== 2) {CHANIS0_5UnderweightM=cl_CHANIS0_5UnderweightM.getRawValue();  } else { CHANIS0_5UnderweightM = "0";}}
XSSFCell cl_CHANIS0_5UnderweightT = worksheet.getRow(47).getCell((short) 16); if ( cl_CHANIS0_5UnderweightT!=null){ if ( cl_CHANIS0_5UnderweightT.getCellType()==0) {CHANIS0_5UnderweightT = "" + (int) cl_CHANIS0_5UnderweightT.getNumericCellValue();} else if (cl_CHANIS0_5UnderweightT.getCellType()==1) {CHANIS0_5UnderweightT=cl_CHANIS0_5UnderweightT.getStringCellValue();}  else if (cl_CHANIS0_5UnderweightT.getCellType()== 2) {CHANIS0_5UnderweightT=cl_CHANIS0_5UnderweightT.getRawValue();  } else { CHANIS0_5UnderweightT = "0";}}
XSSFCell cl_CHANIS0_5sevUnderweightF = worksheet.getRow(48).getCell((short) 14); if ( cl_CHANIS0_5sevUnderweightF!=null){ if ( cl_CHANIS0_5sevUnderweightF.getCellType()==0) {CHANIS0_5sevUnderweightF = "" + (int) cl_CHANIS0_5sevUnderweightF.getNumericCellValue();} else if (cl_CHANIS0_5sevUnderweightF.getCellType()==1) {CHANIS0_5sevUnderweightF=cl_CHANIS0_5sevUnderweightF.getStringCellValue();}  else if (cl_CHANIS0_5sevUnderweightF.getCellType()== 2) {CHANIS0_5sevUnderweightF=cl_CHANIS0_5sevUnderweightF.getRawValue();  } else { CHANIS0_5sevUnderweightF = "0";}}
XSSFCell cl_CHANIS0_5sevUnderweightM = worksheet.getRow(48).getCell((short) 15); if ( cl_CHANIS0_5sevUnderweightM!=null){ if ( cl_CHANIS0_5sevUnderweightM.getCellType()==0) {CHANIS0_5sevUnderweightM = "" + (int) cl_CHANIS0_5sevUnderweightM.getNumericCellValue();} else if (cl_CHANIS0_5sevUnderweightM.getCellType()==1) {CHANIS0_5sevUnderweightM=cl_CHANIS0_5sevUnderweightM.getStringCellValue();}  else if (cl_CHANIS0_5sevUnderweightM.getCellType()== 2) {CHANIS0_5sevUnderweightM=cl_CHANIS0_5sevUnderweightM.getRawValue();  } else { CHANIS0_5sevUnderweightM = "0";}}
XSSFCell cl_CHANIS0_5sevUnderweightT = worksheet.getRow(48).getCell((short) 16); if ( cl_CHANIS0_5sevUnderweightT!=null){ if ( cl_CHANIS0_5sevUnderweightT.getCellType()==0) {CHANIS0_5sevUnderweightT = "" + (int) cl_CHANIS0_5sevUnderweightT.getNumericCellValue();} else if (cl_CHANIS0_5sevUnderweightT.getCellType()==1) {CHANIS0_5sevUnderweightT=cl_CHANIS0_5sevUnderweightT.getStringCellValue();}  else if (cl_CHANIS0_5sevUnderweightT.getCellType()== 2) {CHANIS0_5sevUnderweightT=cl_CHANIS0_5sevUnderweightT.getRawValue();  } else { CHANIS0_5sevUnderweightT = "0";}}
XSSFCell cl_CHANIS0_5OverweightF = worksheet.getRow(49).getCell((short) 14); if ( cl_CHANIS0_5OverweightF!=null){ if ( cl_CHANIS0_5OverweightF.getCellType()==0) {CHANIS0_5OverweightF = "" + (int) cl_CHANIS0_5OverweightF.getNumericCellValue();} else if (cl_CHANIS0_5OverweightF.getCellType()==1) {CHANIS0_5OverweightF=cl_CHANIS0_5OverweightF.getStringCellValue();}  else if (cl_CHANIS0_5OverweightF.getCellType()== 2) {CHANIS0_5OverweightF=cl_CHANIS0_5OverweightF.getRawValue();  } else { CHANIS0_5OverweightF = "0";}}
XSSFCell cl_CHANIS0_5OverweightM = worksheet.getRow(49).getCell((short) 15); if ( cl_CHANIS0_5OverweightM!=null){ if ( cl_CHANIS0_5OverweightM.getCellType()==0) {CHANIS0_5OverweightM = "" + (int) cl_CHANIS0_5OverweightM.getNumericCellValue();} else if (cl_CHANIS0_5OverweightM.getCellType()==1) {CHANIS0_5OverweightM=cl_CHANIS0_5OverweightM.getStringCellValue();}  else if (cl_CHANIS0_5OverweightM.getCellType()== 2) {CHANIS0_5OverweightM=cl_CHANIS0_5OverweightM.getRawValue();  } else { CHANIS0_5OverweightM = "0";}}
XSSFCell cl_CHANIS0_5OverweightT = worksheet.getRow(49).getCell((short) 16); if ( cl_CHANIS0_5OverweightT!=null){ if ( cl_CHANIS0_5OverweightT.getCellType()==0) {CHANIS0_5OverweightT = "" + (int) cl_CHANIS0_5OverweightT.getNumericCellValue();} else if (cl_CHANIS0_5OverweightT.getCellType()==1) {CHANIS0_5OverweightT=cl_CHANIS0_5OverweightT.getStringCellValue();}  else if (cl_CHANIS0_5OverweightT.getCellType()== 2) {CHANIS0_5OverweightT=cl_CHANIS0_5OverweightT.getRawValue();  } else { CHANIS0_5OverweightT = "0";}}
XSSFCell cl_CHANIS0_5ObeseF = worksheet.getRow(50).getCell((short) 14); if ( cl_CHANIS0_5ObeseF!=null){ if ( cl_CHANIS0_5ObeseF.getCellType()==0) {CHANIS0_5ObeseF = "" + (int) cl_CHANIS0_5ObeseF.getNumericCellValue();} else if (cl_CHANIS0_5ObeseF.getCellType()==1) {CHANIS0_5ObeseF=cl_CHANIS0_5ObeseF.getStringCellValue();}  else if (cl_CHANIS0_5ObeseF.getCellType()== 2) {CHANIS0_5ObeseF=cl_CHANIS0_5ObeseF.getRawValue();  } else { CHANIS0_5ObeseF = "0";}}
XSSFCell cl_CHANIS0_5ObeseM = worksheet.getRow(50).getCell((short) 15); if ( cl_CHANIS0_5ObeseM!=null){ if ( cl_CHANIS0_5ObeseM.getCellType()==0) {CHANIS0_5ObeseM = "" + (int) cl_CHANIS0_5ObeseM.getNumericCellValue();} else if (cl_CHANIS0_5ObeseM.getCellType()==1) {CHANIS0_5ObeseM=cl_CHANIS0_5ObeseM.getStringCellValue();}  else if (cl_CHANIS0_5ObeseM.getCellType()== 2) {CHANIS0_5ObeseM=cl_CHANIS0_5ObeseM.getRawValue();  } else { CHANIS0_5ObeseM = "0";}}
XSSFCell cl_CHANIS0_5ObeseT = worksheet.getRow(50).getCell((short) 16); if ( cl_CHANIS0_5ObeseT!=null){ if ( cl_CHANIS0_5ObeseT.getCellType()==0) {CHANIS0_5ObeseT = "" + (int) cl_CHANIS0_5ObeseT.getNumericCellValue();} else if (cl_CHANIS0_5ObeseT.getCellType()==1) {CHANIS0_5ObeseT=cl_CHANIS0_5ObeseT.getStringCellValue();}  else if (cl_CHANIS0_5ObeseT.getCellType()== 2) {CHANIS0_5ObeseT=cl_CHANIS0_5ObeseT.getRawValue();  } else { CHANIS0_5ObeseT = "0";}}
XSSFCell cl_CHANIS0_5TWF = worksheet.getRow(51).getCell((short) 14); if ( cl_CHANIS0_5TWF!=null){ if ( cl_CHANIS0_5TWF.getCellType()==0) {CHANIS0_5TWF = "" + (int) cl_CHANIS0_5TWF.getNumericCellValue();} else if (cl_CHANIS0_5TWF.getCellType()==1) {CHANIS0_5TWF=cl_CHANIS0_5TWF.getStringCellValue();}  else if (cl_CHANIS0_5TWF.getCellType()== 2) {CHANIS0_5TWF=cl_CHANIS0_5TWF.getRawValue();  } else { CHANIS0_5TWF = "0";}}
XSSFCell cl_CHANIS0_5TWM = worksheet.getRow(51).getCell((short) 15); if ( cl_CHANIS0_5TWM!=null){ if ( cl_CHANIS0_5TWM.getCellType()==0) {CHANIS0_5TWM = "" + (int) cl_CHANIS0_5TWM.getNumericCellValue();} else if (cl_CHANIS0_5TWM.getCellType()==1) {CHANIS0_5TWM=cl_CHANIS0_5TWM.getStringCellValue();}  else if (cl_CHANIS0_5TWM.getCellType()== 2) {CHANIS0_5TWM=cl_CHANIS0_5TWM.getRawValue();  } else { CHANIS0_5TWM = "0";}}
XSSFCell cl_CHANIS0_5TW = worksheet.getRow(51).getCell((short) 16); if ( cl_CHANIS0_5TW!=null){ if ( cl_CHANIS0_5TW.getCellType()==0) {CHANIS0_5TW = "" + (int) cl_CHANIS0_5TW.getNumericCellValue();} else if (cl_CHANIS0_5TW.getCellType()==1) {CHANIS0_5TW=cl_CHANIS0_5TW.getStringCellValue();}  else if (cl_CHANIS0_5TW.getCellType()== 2) {CHANIS0_5TW=cl_CHANIS0_5TW.getRawValue();  } else { CHANIS0_5TW = "0";}}
XSSFCell cl_CHANIS6_23NormalweightF = worksheet.getRow(52).getCell((short) 14); if ( cl_CHANIS6_23NormalweightF!=null){ if ( cl_CHANIS6_23NormalweightF.getCellType()==0) {CHANIS6_23NormalweightF = "" + (int) cl_CHANIS6_23NormalweightF.getNumericCellValue();} else if (cl_CHANIS6_23NormalweightF.getCellType()==1) {CHANIS6_23NormalweightF=cl_CHANIS6_23NormalweightF.getStringCellValue();}  else if (cl_CHANIS6_23NormalweightF.getCellType()== 2) {CHANIS6_23NormalweightF=cl_CHANIS6_23NormalweightF.getRawValue();  } else { CHANIS6_23NormalweightF = "0";}}
XSSFCell cl_CHANIS6_23NormalweightM = worksheet.getRow(52).getCell((short) 15); if ( cl_CHANIS6_23NormalweightM!=null){ if ( cl_CHANIS6_23NormalweightM.getCellType()==0) {CHANIS6_23NormalweightM = "" + (int) cl_CHANIS6_23NormalweightM.getNumericCellValue();} else if (cl_CHANIS6_23NormalweightM.getCellType()==1) {CHANIS6_23NormalweightM=cl_CHANIS6_23NormalweightM.getStringCellValue();}  else if (cl_CHANIS6_23NormalweightM.getCellType()== 2) {CHANIS6_23NormalweightM=cl_CHANIS6_23NormalweightM.getRawValue();  } else { CHANIS6_23NormalweightM = "0";}}
XSSFCell cl_CHANIS6_23NormalweightT = worksheet.getRow(52).getCell((short) 16); if ( cl_CHANIS6_23NormalweightT!=null){ if ( cl_CHANIS6_23NormalweightT.getCellType()==0) {CHANIS6_23NormalweightT = "" + (int) cl_CHANIS6_23NormalweightT.getNumericCellValue();} else if (cl_CHANIS6_23NormalweightT.getCellType()==1) {CHANIS6_23NormalweightT=cl_CHANIS6_23NormalweightT.getStringCellValue();}  else if (cl_CHANIS6_23NormalweightT.getCellType()== 2) {CHANIS6_23NormalweightT=cl_CHANIS6_23NormalweightT.getRawValue();  } else { CHANIS6_23NormalweightT = "0";}}
XSSFCell cl_CHANIS6_23UnderweightF = worksheet.getRow(53).getCell((short) 14); if ( cl_CHANIS6_23UnderweightF!=null){ if ( cl_CHANIS6_23UnderweightF.getCellType()==0) {CHANIS6_23UnderweightF = "" + (int) cl_CHANIS6_23UnderweightF.getNumericCellValue();} else if (cl_CHANIS6_23UnderweightF.getCellType()==1) {CHANIS6_23UnderweightF=cl_CHANIS6_23UnderweightF.getStringCellValue();}  else if (cl_CHANIS6_23UnderweightF.getCellType()== 2) {CHANIS6_23UnderweightF=cl_CHANIS6_23UnderweightF.getRawValue();  } else { CHANIS6_23UnderweightF = "0";}}
XSSFCell cl_CHANIS6_23UnderweightM = worksheet.getRow(53).getCell((short) 15); if ( cl_CHANIS6_23UnderweightM!=null){ if ( cl_CHANIS6_23UnderweightM.getCellType()==0) {CHANIS6_23UnderweightM = "" + (int) cl_CHANIS6_23UnderweightM.getNumericCellValue();} else if (cl_CHANIS6_23UnderweightM.getCellType()==1) {CHANIS6_23UnderweightM=cl_CHANIS6_23UnderweightM.getStringCellValue();}  else if (cl_CHANIS6_23UnderweightM.getCellType()== 2) {CHANIS6_23UnderweightM=cl_CHANIS6_23UnderweightM.getRawValue();  } else { CHANIS6_23UnderweightM = "0";}}
XSSFCell cl_CHANIS6_23UnderweightT = worksheet.getRow(53).getCell((short) 16); if ( cl_CHANIS6_23UnderweightT!=null){ if ( cl_CHANIS6_23UnderweightT.getCellType()==0) {CHANIS6_23UnderweightT = "" + (int) cl_CHANIS6_23UnderweightT.getNumericCellValue();} else if (cl_CHANIS6_23UnderweightT.getCellType()==1) {CHANIS6_23UnderweightT=cl_CHANIS6_23UnderweightT.getStringCellValue();}  else if (cl_CHANIS6_23UnderweightT.getCellType()== 2) {CHANIS6_23UnderweightT=cl_CHANIS6_23UnderweightT.getRawValue();  } else { CHANIS6_23UnderweightT = "0";}}
XSSFCell cl_CHANIS6_23sevUnderweightF = worksheet.getRow(54).getCell((short) 14); if ( cl_CHANIS6_23sevUnderweightF!=null){ if ( cl_CHANIS6_23sevUnderweightF.getCellType()==0) {CHANIS6_23sevUnderweightF = "" + (int) cl_CHANIS6_23sevUnderweightF.getNumericCellValue();} else if (cl_CHANIS6_23sevUnderweightF.getCellType()==1) {CHANIS6_23sevUnderweightF=cl_CHANIS6_23sevUnderweightF.getStringCellValue();}  else if (cl_CHANIS6_23sevUnderweightF.getCellType()== 2) {CHANIS6_23sevUnderweightF=cl_CHANIS6_23sevUnderweightF.getRawValue();  } else { CHANIS6_23sevUnderweightF = "0";}}
XSSFCell cl_CHANIS6_23sevUnderweightM = worksheet.getRow(54).getCell((short) 15); if ( cl_CHANIS6_23sevUnderweightM!=null){ if ( cl_CHANIS6_23sevUnderweightM.getCellType()==0) {CHANIS6_23sevUnderweightM = "" + (int) cl_CHANIS6_23sevUnderweightM.getNumericCellValue();} else if (cl_CHANIS6_23sevUnderweightM.getCellType()==1) {CHANIS6_23sevUnderweightM=cl_CHANIS6_23sevUnderweightM.getStringCellValue();}  else if (cl_CHANIS6_23sevUnderweightM.getCellType()== 2) {CHANIS6_23sevUnderweightM=cl_CHANIS6_23sevUnderweightM.getRawValue();  } else { CHANIS6_23sevUnderweightM = "0";}}
XSSFCell cl_CHANIS6_23sevUnderweightT = worksheet.getRow(54).getCell((short) 16); if ( cl_CHANIS6_23sevUnderweightT!=null){ if ( cl_CHANIS6_23sevUnderweightT.getCellType()==0) {CHANIS6_23sevUnderweightT = "" + (int) cl_CHANIS6_23sevUnderweightT.getNumericCellValue();} else if (cl_CHANIS6_23sevUnderweightT.getCellType()==1) {CHANIS6_23sevUnderweightT=cl_CHANIS6_23sevUnderweightT.getStringCellValue();}  else if (cl_CHANIS6_23sevUnderweightT.getCellType()== 2) {CHANIS6_23sevUnderweightT=cl_CHANIS6_23sevUnderweightT.getRawValue();  } else { CHANIS6_23sevUnderweightT = "0";}}
XSSFCell cl_CHANIS6_23OverweightF = worksheet.getRow(55).getCell((short) 14); if ( cl_CHANIS6_23OverweightF!=null){ if ( cl_CHANIS6_23OverweightF.getCellType()==0) {CHANIS6_23OverweightF = "" + (int) cl_CHANIS6_23OverweightF.getNumericCellValue();} else if (cl_CHANIS6_23OverweightF.getCellType()==1) {CHANIS6_23OverweightF=cl_CHANIS6_23OverweightF.getStringCellValue();}  else if (cl_CHANIS6_23OverweightF.getCellType()== 2) {CHANIS6_23OverweightF=cl_CHANIS6_23OverweightF.getRawValue();  } else { CHANIS6_23OverweightF = "0";}}
XSSFCell cl_CHANIS6_23OverweightM = worksheet.getRow(55).getCell((short) 15); if ( cl_CHANIS6_23OverweightM!=null){ if ( cl_CHANIS6_23OverweightM.getCellType()==0) {CHANIS6_23OverweightM = "" + (int) cl_CHANIS6_23OverweightM.getNumericCellValue();} else if (cl_CHANIS6_23OverweightM.getCellType()==1) {CHANIS6_23OverweightM=cl_CHANIS6_23OverweightM.getStringCellValue();}  else if (cl_CHANIS6_23OverweightM.getCellType()== 2) {CHANIS6_23OverweightM=cl_CHANIS6_23OverweightM.getRawValue();  } else { CHANIS6_23OverweightM = "0";}}
XSSFCell cl_CHANIS6_23OverweightT = worksheet.getRow(55).getCell((short) 16); if ( cl_CHANIS6_23OverweightT!=null){ if ( cl_CHANIS6_23OverweightT.getCellType()==0) {CHANIS6_23OverweightT = "" + (int) cl_CHANIS6_23OverweightT.getNumericCellValue();} else if (cl_CHANIS6_23OverweightT.getCellType()==1) {CHANIS6_23OverweightT=cl_CHANIS6_23OverweightT.getStringCellValue();}  else if (cl_CHANIS6_23OverweightT.getCellType()== 2) {CHANIS6_23OverweightT=cl_CHANIS6_23OverweightT.getRawValue();  } else { CHANIS6_23OverweightT = "0";}}
XSSFCell cl_CHANIS6_23ObeseF = worksheet.getRow(56).getCell((short) 14); if ( cl_CHANIS6_23ObeseF!=null){ if ( cl_CHANIS6_23ObeseF.getCellType()==0) {CHANIS6_23ObeseF = "" + (int) cl_CHANIS6_23ObeseF.getNumericCellValue();} else if (cl_CHANIS6_23ObeseF.getCellType()==1) {CHANIS6_23ObeseF=cl_CHANIS6_23ObeseF.getStringCellValue();}  else if (cl_CHANIS6_23ObeseF.getCellType()== 2) {CHANIS6_23ObeseF=cl_CHANIS6_23ObeseF.getRawValue();  } else { CHANIS6_23ObeseF = "0";}}
XSSFCell cl_CHANIS6_23ObeseM = worksheet.getRow(56).getCell((short) 15); if ( cl_CHANIS6_23ObeseM!=null){ if ( cl_CHANIS6_23ObeseM.getCellType()==0) {CHANIS6_23ObeseM = "" + (int) cl_CHANIS6_23ObeseM.getNumericCellValue();} else if (cl_CHANIS6_23ObeseM.getCellType()==1) {CHANIS6_23ObeseM=cl_CHANIS6_23ObeseM.getStringCellValue();}  else if (cl_CHANIS6_23ObeseM.getCellType()== 2) {CHANIS6_23ObeseM=cl_CHANIS6_23ObeseM.getRawValue();  } else { CHANIS6_23ObeseM = "0";}}
XSSFCell cl_CHANIS6_23ObeseT = worksheet.getRow(56).getCell((short) 16); if ( cl_CHANIS6_23ObeseT!=null){ if ( cl_CHANIS6_23ObeseT.getCellType()==0) {CHANIS6_23ObeseT = "" + (int) cl_CHANIS6_23ObeseT.getNumericCellValue();} else if (cl_CHANIS6_23ObeseT.getCellType()==1) {CHANIS6_23ObeseT=cl_CHANIS6_23ObeseT.getStringCellValue();}  else if (cl_CHANIS6_23ObeseT.getCellType()== 2) {CHANIS6_23ObeseT=cl_CHANIS6_23ObeseT.getRawValue();  } else { CHANIS6_23ObeseT = "0";}}
XSSFCell cl_CHANIS6_23TWF = worksheet.getRow(57).getCell((short) 14); if ( cl_CHANIS6_23TWF!=null){ if ( cl_CHANIS6_23TWF.getCellType()==0) {CHANIS6_23TWF = "" + (int) cl_CHANIS6_23TWF.getNumericCellValue();} else if (cl_CHANIS6_23TWF.getCellType()==1) {CHANIS6_23TWF=cl_CHANIS6_23TWF.getStringCellValue();}  else if (cl_CHANIS6_23TWF.getCellType()== 2) {CHANIS6_23TWF=cl_CHANIS6_23TWF.getRawValue();  } else { CHANIS6_23TWF = "0";}}
XSSFCell cl_CHANIS6_23TWM = worksheet.getRow(57).getCell((short) 15); if ( cl_CHANIS6_23TWM!=null){ if ( cl_CHANIS6_23TWM.getCellType()==0) {CHANIS6_23TWM = "" + (int) cl_CHANIS6_23TWM.getNumericCellValue();} else if (cl_CHANIS6_23TWM.getCellType()==1) {CHANIS6_23TWM=cl_CHANIS6_23TWM.getStringCellValue();}  else if (cl_CHANIS6_23TWM.getCellType()== 2) {CHANIS6_23TWM=cl_CHANIS6_23TWM.getRawValue();  } else { CHANIS6_23TWM = "0";}}
XSSFCell cl_CHANIS6_23TW = worksheet.getRow(57).getCell((short) 16); if ( cl_CHANIS6_23TW!=null){ if ( cl_CHANIS6_23TW.getCellType()==0) {CHANIS6_23TW = "" + (int) cl_CHANIS6_23TW.getNumericCellValue();} else if (cl_CHANIS6_23TW.getCellType()==1) {CHANIS6_23TW=cl_CHANIS6_23TW.getStringCellValue();}  else if (cl_CHANIS6_23TW.getCellType()== 2) {CHANIS6_23TW=cl_CHANIS6_23TW.getRawValue();  } else { CHANIS6_23TW = "0";}}
XSSFCell cl_CHANIS24_59NormalweightF = worksheet.getRow(58).getCell((short) 14); if ( cl_CHANIS24_59NormalweightF!=null){ if ( cl_CHANIS24_59NormalweightF.getCellType()==0) {CHANIS24_59NormalweightF = "" + (int) cl_CHANIS24_59NormalweightF.getNumericCellValue();} else if (cl_CHANIS24_59NormalweightF.getCellType()==1) {CHANIS24_59NormalweightF=cl_CHANIS24_59NormalweightF.getStringCellValue();}  else if (cl_CHANIS24_59NormalweightF.getCellType()== 2) {CHANIS24_59NormalweightF=cl_CHANIS24_59NormalweightF.getRawValue();  } else { CHANIS24_59NormalweightF = "0";}}
XSSFCell cl_CHANIS24_59NormalweightM = worksheet.getRow(58).getCell((short) 15); if ( cl_CHANIS24_59NormalweightM!=null){ if ( cl_CHANIS24_59NormalweightM.getCellType()==0) {CHANIS24_59NormalweightM = "" + (int) cl_CHANIS24_59NormalweightM.getNumericCellValue();} else if (cl_CHANIS24_59NormalweightM.getCellType()==1) {CHANIS24_59NormalweightM=cl_CHANIS24_59NormalweightM.getStringCellValue();}  else if (cl_CHANIS24_59NormalweightM.getCellType()== 2) {CHANIS24_59NormalweightM=cl_CHANIS24_59NormalweightM.getRawValue();  } else { CHANIS24_59NormalweightM = "0";}}
XSSFCell cl_CHANIS24_59NormalweightT = worksheet.getRow(58).getCell((short) 16); if ( cl_CHANIS24_59NormalweightT!=null){ if ( cl_CHANIS24_59NormalweightT.getCellType()==0) {CHANIS24_59NormalweightT = "" + (int) cl_CHANIS24_59NormalweightT.getNumericCellValue();} else if (cl_CHANIS24_59NormalweightT.getCellType()==1) {CHANIS24_59NormalweightT=cl_CHANIS24_59NormalweightT.getStringCellValue();}  else if (cl_CHANIS24_59NormalweightT.getCellType()== 2) {CHANIS24_59NormalweightT=cl_CHANIS24_59NormalweightT.getRawValue();  } else { CHANIS24_59NormalweightT = "0";}}
XSSFCell cl_CHANIS24_59UnderweightF = worksheet.getRow(59).getCell((short) 14); if ( cl_CHANIS24_59UnderweightF!=null){ if ( cl_CHANIS24_59UnderweightF.getCellType()==0) {CHANIS24_59UnderweightF = "" + (int) cl_CHANIS24_59UnderweightF.getNumericCellValue();} else if (cl_CHANIS24_59UnderweightF.getCellType()==1) {CHANIS24_59UnderweightF=cl_CHANIS24_59UnderweightF.getStringCellValue();}  else if (cl_CHANIS24_59UnderweightF.getCellType()== 2) {CHANIS24_59UnderweightF=cl_CHANIS24_59UnderweightF.getRawValue();  } else { CHANIS24_59UnderweightF = "0";}}
XSSFCell cl_CHANIS24_59UnderweightM = worksheet.getRow(59).getCell((short) 15); if ( cl_CHANIS24_59UnderweightM!=null){ if ( cl_CHANIS24_59UnderweightM.getCellType()==0) {CHANIS24_59UnderweightM = "" + (int) cl_CHANIS24_59UnderweightM.getNumericCellValue();} else if (cl_CHANIS24_59UnderweightM.getCellType()==1) {CHANIS24_59UnderweightM=cl_CHANIS24_59UnderweightM.getStringCellValue();}  else if (cl_CHANIS24_59UnderweightM.getCellType()== 2) {CHANIS24_59UnderweightM=cl_CHANIS24_59UnderweightM.getRawValue();  } else { CHANIS24_59UnderweightM = "0";}}
XSSFCell cl_CHANIS24_59UnderweightT = worksheet.getRow(59).getCell((short) 16); if ( cl_CHANIS24_59UnderweightT!=null){ if ( cl_CHANIS24_59UnderweightT.getCellType()==0) {CHANIS24_59UnderweightT = "" + (int) cl_CHANIS24_59UnderweightT.getNumericCellValue();} else if (cl_CHANIS24_59UnderweightT.getCellType()==1) {CHANIS24_59UnderweightT=cl_CHANIS24_59UnderweightT.getStringCellValue();}  else if (cl_CHANIS24_59UnderweightT.getCellType()== 2) {CHANIS24_59UnderweightT=cl_CHANIS24_59UnderweightT.getRawValue();  } else { CHANIS24_59UnderweightT = "0";}}
XSSFCell cl_CHANIS24_59sevUnderweightF = worksheet.getRow(60).getCell((short) 14); if ( cl_CHANIS24_59sevUnderweightF!=null){ if ( cl_CHANIS24_59sevUnderweightF.getCellType()==0) {CHANIS24_59sevUnderweightF = "" + (int) cl_CHANIS24_59sevUnderweightF.getNumericCellValue();} else if (cl_CHANIS24_59sevUnderweightF.getCellType()==1) {CHANIS24_59sevUnderweightF=cl_CHANIS24_59sevUnderweightF.getStringCellValue();}  else if (cl_CHANIS24_59sevUnderweightF.getCellType()== 2) {CHANIS24_59sevUnderweightF=cl_CHANIS24_59sevUnderweightF.getRawValue();  } else { CHANIS24_59sevUnderweightF = "0";}}
XSSFCell cl_CHANIS24_59sevUnderweightM = worksheet.getRow(60).getCell((short) 15); if ( cl_CHANIS24_59sevUnderweightM!=null){ if ( cl_CHANIS24_59sevUnderweightM.getCellType()==0) {CHANIS24_59sevUnderweightM = "" + (int) cl_CHANIS24_59sevUnderweightM.getNumericCellValue();} else if (cl_CHANIS24_59sevUnderweightM.getCellType()==1) {CHANIS24_59sevUnderweightM=cl_CHANIS24_59sevUnderweightM.getStringCellValue();}  else if (cl_CHANIS24_59sevUnderweightM.getCellType()== 2) {CHANIS24_59sevUnderweightM=cl_CHANIS24_59sevUnderweightM.getRawValue();  } else { CHANIS24_59sevUnderweightM = "0";}}
XSSFCell cl_CHANIS24_59sevUnderweightT = worksheet.getRow(60).getCell((short) 16); if ( cl_CHANIS24_59sevUnderweightT!=null){ if ( cl_CHANIS24_59sevUnderweightT.getCellType()==0) {CHANIS24_59sevUnderweightT = "" + (int) cl_CHANIS24_59sevUnderweightT.getNumericCellValue();} else if (cl_CHANIS24_59sevUnderweightT.getCellType()==1) {CHANIS24_59sevUnderweightT=cl_CHANIS24_59sevUnderweightT.getStringCellValue();}  else if (cl_CHANIS24_59sevUnderweightT.getCellType()== 2) {CHANIS24_59sevUnderweightT=cl_CHANIS24_59sevUnderweightT.getRawValue();  } else { CHANIS24_59sevUnderweightT = "0";}}
XSSFCell cl_CHANIS24_59OverweightF = worksheet.getRow(61).getCell((short) 14); if ( cl_CHANIS24_59OverweightF!=null){ if ( cl_CHANIS24_59OverweightF.getCellType()==0) {CHANIS24_59OverweightF = "" + (int) cl_CHANIS24_59OverweightF.getNumericCellValue();} else if (cl_CHANIS24_59OverweightF.getCellType()==1) {CHANIS24_59OverweightF=cl_CHANIS24_59OverweightF.getStringCellValue();}  else if (cl_CHANIS24_59OverweightF.getCellType()== 2) {CHANIS24_59OverweightF=cl_CHANIS24_59OverweightF.getRawValue();  } else { CHANIS24_59OverweightF = "0";}}
XSSFCell cl_CHANIS24_59OverweightM = worksheet.getRow(61).getCell((short) 15); if ( cl_CHANIS24_59OverweightM!=null){ if ( cl_CHANIS24_59OverweightM.getCellType()==0) {CHANIS24_59OverweightM = "" + (int) cl_CHANIS24_59OverweightM.getNumericCellValue();} else if (cl_CHANIS24_59OverweightM.getCellType()==1) {CHANIS24_59OverweightM=cl_CHANIS24_59OverweightM.getStringCellValue();}  else if (cl_CHANIS24_59OverweightM.getCellType()== 2) {CHANIS24_59OverweightM=cl_CHANIS24_59OverweightM.getRawValue();  } else { CHANIS24_59OverweightM = "0";}}
XSSFCell cl_CHANIS24_59OverweightT = worksheet.getRow(61).getCell((short) 16); if ( cl_CHANIS24_59OverweightT!=null){ if ( cl_CHANIS24_59OverweightT.getCellType()==0) {CHANIS24_59OverweightT = "" + (int) cl_CHANIS24_59OverweightT.getNumericCellValue();} else if (cl_CHANIS24_59OverweightT.getCellType()==1) {CHANIS24_59OverweightT=cl_CHANIS24_59OverweightT.getStringCellValue();}  else if (cl_CHANIS24_59OverweightT.getCellType()== 2) {CHANIS24_59OverweightT=cl_CHANIS24_59OverweightT.getRawValue();  } else { CHANIS24_59OverweightT = "0";}}
XSSFCell cl_CHANIS24_59ObeseF = worksheet.getRow(62).getCell((short) 14); if ( cl_CHANIS24_59ObeseF!=null){ if ( cl_CHANIS24_59ObeseF.getCellType()==0) {CHANIS24_59ObeseF = "" + (int) cl_CHANIS24_59ObeseF.getNumericCellValue();} else if (cl_CHANIS24_59ObeseF.getCellType()==1) {CHANIS24_59ObeseF=cl_CHANIS24_59ObeseF.getStringCellValue();}  else if (cl_CHANIS24_59ObeseF.getCellType()== 2) {CHANIS24_59ObeseF=cl_CHANIS24_59ObeseF.getRawValue();  } else { CHANIS24_59ObeseF = "0";}}
XSSFCell cl_CHANIS24_59ObeseM = worksheet.getRow(62).getCell((short) 15); if ( cl_CHANIS24_59ObeseM!=null){ if ( cl_CHANIS24_59ObeseM.getCellType()==0) {CHANIS24_59ObeseM = "" + (int) cl_CHANIS24_59ObeseM.getNumericCellValue();} else if (cl_CHANIS24_59ObeseM.getCellType()==1) {CHANIS24_59ObeseM=cl_CHANIS24_59ObeseM.getStringCellValue();}  else if (cl_CHANIS24_59ObeseM.getCellType()== 2) {CHANIS24_59ObeseM=cl_CHANIS24_59ObeseM.getRawValue();  } else { CHANIS24_59ObeseM = "0";}}
XSSFCell cl_CHANIS24_59ObeseT = worksheet.getRow(62).getCell((short) 16); if ( cl_CHANIS24_59ObeseT!=null){ if ( cl_CHANIS24_59ObeseT.getCellType()==0) {CHANIS24_59ObeseT = "" + (int) cl_CHANIS24_59ObeseT.getNumericCellValue();} else if (cl_CHANIS24_59ObeseT.getCellType()==1) {CHANIS24_59ObeseT=cl_CHANIS24_59ObeseT.getStringCellValue();}  else if (cl_CHANIS24_59ObeseT.getCellType()== 2) {CHANIS24_59ObeseT=cl_CHANIS24_59ObeseT.getRawValue();  } else { CHANIS24_59ObeseT = "0";}}
XSSFCell cl_CHANIS24_59TWF = worksheet.getRow(63).getCell((short) 14); if ( cl_CHANIS24_59TWF!=null){ if ( cl_CHANIS24_59TWF.getCellType()==0) {CHANIS24_59TWF = "" + (int) cl_CHANIS24_59TWF.getNumericCellValue();} else if (cl_CHANIS24_59TWF.getCellType()==1) {CHANIS24_59TWF=cl_CHANIS24_59TWF.getStringCellValue();}  else if (cl_CHANIS24_59TWF.getCellType()== 2) {CHANIS24_59TWF=cl_CHANIS24_59TWF.getRawValue();  } else { CHANIS24_59TWF = "0";}}
XSSFCell cl_CHANIS24_59TWM = worksheet.getRow(63).getCell((short) 15); if ( cl_CHANIS24_59TWM!=null){ if ( cl_CHANIS24_59TWM.getCellType()==0) {CHANIS24_59TWM = "" + (int) cl_CHANIS24_59TWM.getNumericCellValue();} else if (cl_CHANIS24_59TWM.getCellType()==1) {CHANIS24_59TWM=cl_CHANIS24_59TWM.getStringCellValue();}  else if (cl_CHANIS24_59TWM.getCellType()== 2) {CHANIS24_59TWM=cl_CHANIS24_59TWM.getRawValue();  } else { CHANIS24_59TWM = "0";}}
XSSFCell cl_CHANIS24_59TW = worksheet.getRow(63).getCell((short) 16); if ( cl_CHANIS24_59TW!=null){ if ( cl_CHANIS24_59TW.getCellType()==0) {CHANIS24_59TW = "" + (int) cl_CHANIS24_59TW.getNumericCellValue();} else if (cl_CHANIS24_59TW.getCellType()==1) {CHANIS24_59TW=cl_CHANIS24_59TW.getStringCellValue();}  else if (cl_CHANIS24_59TW.getCellType()== 2) {CHANIS24_59TW=cl_CHANIS24_59TW.getRawValue();  } else { CHANIS24_59TW = "0";}}
XSSFCell cl_CHANISMUACNormalF = worksheet.getRow(64).getCell((short) 14); if ( cl_CHANISMUACNormalF!=null){ if ( cl_CHANISMUACNormalF.getCellType()==0) {CHANISMUACNormalF = "" + (int) cl_CHANISMUACNormalF.getNumericCellValue();} else if (cl_CHANISMUACNormalF.getCellType()==1) {CHANISMUACNormalF=cl_CHANISMUACNormalF.getStringCellValue();}  else if (cl_CHANISMUACNormalF.getCellType()== 2) {CHANISMUACNormalF=cl_CHANISMUACNormalF.getRawValue();  } else { CHANISMUACNormalF = "0";}}
XSSFCell cl_CHANISMUACNormalM = worksheet.getRow(64).getCell((short) 15); if ( cl_CHANISMUACNormalM!=null){ if ( cl_CHANISMUACNormalM.getCellType()==0) {CHANISMUACNormalM = "" + (int) cl_CHANISMUACNormalM.getNumericCellValue();} else if (cl_CHANISMUACNormalM.getCellType()==1) {CHANISMUACNormalM=cl_CHANISMUACNormalM.getStringCellValue();}  else if (cl_CHANISMUACNormalM.getCellType()== 2) {CHANISMUACNormalM=cl_CHANISMUACNormalM.getRawValue();  } else { CHANISMUACNormalM = "0";}}
XSSFCell cl_CHANISMUACNormalT = worksheet.getRow(64).getCell((short) 16); if ( cl_CHANISMUACNormalT!=null){ if ( cl_CHANISMUACNormalT.getCellType()==0) {CHANISMUACNormalT = "" + (int) cl_CHANISMUACNormalT.getNumericCellValue();} else if (cl_CHANISMUACNormalT.getCellType()==1) {CHANISMUACNormalT=cl_CHANISMUACNormalT.getStringCellValue();}  else if (cl_CHANISMUACNormalT.getCellType()== 2) {CHANISMUACNormalT=cl_CHANISMUACNormalT.getRawValue();  } else { CHANISMUACNormalT = "0";}}
XSSFCell cl_CHANISMUACModerateF = worksheet.getRow(65).getCell((short) 14); if ( cl_CHANISMUACModerateF!=null){ if ( cl_CHANISMUACModerateF.getCellType()==0) {CHANISMUACModerateF = "" + (int) cl_CHANISMUACModerateF.getNumericCellValue();} else if (cl_CHANISMUACModerateF.getCellType()==1) {CHANISMUACModerateF=cl_CHANISMUACModerateF.getStringCellValue();}  else if (cl_CHANISMUACModerateF.getCellType()== 2) {CHANISMUACModerateF=cl_CHANISMUACModerateF.getRawValue();  } else { CHANISMUACModerateF = "0";}}
XSSFCell cl_CHANISMUACModerateM = worksheet.getRow(65).getCell((short) 15); if ( cl_CHANISMUACModerateM!=null){ if ( cl_CHANISMUACModerateM.getCellType()==0) {CHANISMUACModerateM = "" + (int) cl_CHANISMUACModerateM.getNumericCellValue();} else if (cl_CHANISMUACModerateM.getCellType()==1) {CHANISMUACModerateM=cl_CHANISMUACModerateM.getStringCellValue();}  else if (cl_CHANISMUACModerateM.getCellType()== 2) {CHANISMUACModerateM=cl_CHANISMUACModerateM.getRawValue();  } else { CHANISMUACModerateM = "0";}}
XSSFCell cl_CHANISMUACModerateT = worksheet.getRow(65).getCell((short) 16); if ( cl_CHANISMUACModerateT!=null){ if ( cl_CHANISMUACModerateT.getCellType()==0) {CHANISMUACModerateT = "" + (int) cl_CHANISMUACModerateT.getNumericCellValue();} else if (cl_CHANISMUACModerateT.getCellType()==1) {CHANISMUACModerateT=cl_CHANISMUACModerateT.getStringCellValue();}  else if (cl_CHANISMUACModerateT.getCellType()== 2) {CHANISMUACModerateT=cl_CHANISMUACModerateT.getRawValue();  } else { CHANISMUACModerateT = "0";}}
XSSFCell cl_CHANISMUACSevereF = worksheet.getRow(66).getCell((short) 14); if ( cl_CHANISMUACSevereF!=null){ if ( cl_CHANISMUACSevereF.getCellType()==0) {CHANISMUACSevereF = "" + (int) cl_CHANISMUACSevereF.getNumericCellValue();} else if (cl_CHANISMUACSevereF.getCellType()==1) {CHANISMUACSevereF=cl_CHANISMUACSevereF.getStringCellValue();}  else if (cl_CHANISMUACSevereF.getCellType()== 2) {CHANISMUACSevereF=cl_CHANISMUACSevereF.getRawValue();  } else { CHANISMUACSevereF = "0";}}
XSSFCell cl_CHANISMUACSevereM = worksheet.getRow(66).getCell((short) 15); if ( cl_CHANISMUACSevereM!=null){ if ( cl_CHANISMUACSevereM.getCellType()==0) {CHANISMUACSevereM = "" + (int) cl_CHANISMUACSevereM.getNumericCellValue();} else if (cl_CHANISMUACSevereM.getCellType()==1) {CHANISMUACSevereM=cl_CHANISMUACSevereM.getStringCellValue();}  else if (cl_CHANISMUACSevereM.getCellType()== 2) {CHANISMUACSevereM=cl_CHANISMUACSevereM.getRawValue();  } else { CHANISMUACSevereM = "0";}}
XSSFCell cl_CHANISMUACSevereT = worksheet.getRow(66).getCell((short) 16); if ( cl_CHANISMUACSevereT!=null){ if ( cl_CHANISMUACSevereT.getCellType()==0) {CHANISMUACSevereT = "" + (int) cl_CHANISMUACSevereT.getNumericCellValue();} else if (cl_CHANISMUACSevereT.getCellType()==1) {CHANISMUACSevereT=cl_CHANISMUACSevereT.getStringCellValue();}  else if (cl_CHANISMUACSevereT.getCellType()== 2) {CHANISMUACSevereT=cl_CHANISMUACSevereT.getRawValue();  } else { CHANISMUACSevereT = "0";}}
XSSFCell cl_CHANISMUACMeasuredF = worksheet.getRow(67).getCell((short) 14); if ( cl_CHANISMUACMeasuredF!=null){ if ( cl_CHANISMUACMeasuredF.getCellType()==0) {CHANISMUACMeasuredF = "" + (int) cl_CHANISMUACMeasuredF.getNumericCellValue();} else if (cl_CHANISMUACMeasuredF.getCellType()==1) {CHANISMUACMeasuredF=cl_CHANISMUACMeasuredF.getStringCellValue();}  else if (cl_CHANISMUACMeasuredF.getCellType()== 2) {CHANISMUACMeasuredF=cl_CHANISMUACMeasuredF.getRawValue();  } else { CHANISMUACMeasuredF = "0";}}
XSSFCell cl_CHANISMUACMeasuredM = worksheet.getRow(67).getCell((short) 15); if ( cl_CHANISMUACMeasuredM!=null){ if ( cl_CHANISMUACMeasuredM.getCellType()==0) {CHANISMUACMeasuredM = "" + (int) cl_CHANISMUACMeasuredM.getNumericCellValue();} else if (cl_CHANISMUACMeasuredM.getCellType()==1) {CHANISMUACMeasuredM=cl_CHANISMUACMeasuredM.getStringCellValue();}  else if (cl_CHANISMUACMeasuredM.getCellType()== 2) {CHANISMUACMeasuredM=cl_CHANISMUACMeasuredM.getRawValue();  } else { CHANISMUACMeasuredM = "0";}}
XSSFCell cl_CHANISMUACMeasuredT = worksheet.getRow(67).getCell((short) 16); if ( cl_CHANISMUACMeasuredT!=null){ if ( cl_CHANISMUACMeasuredT.getCellType()==0) {CHANISMUACMeasuredT = "" + (int) cl_CHANISMUACMeasuredT.getNumericCellValue();} else if (cl_CHANISMUACMeasuredT.getCellType()==1) {CHANISMUACMeasuredT=cl_CHANISMUACMeasuredT.getStringCellValue();}  else if (cl_CHANISMUACMeasuredT.getCellType()== 2) {CHANISMUACMeasuredT=cl_CHANISMUACMeasuredT.getRawValue();  } else { CHANISMUACMeasuredT = "0";}}
XSSFCell cl_CHANIS0_5NormalHeightF = worksheet.getRow(69).getCell((short) 14); if ( cl_CHANIS0_5NormalHeightF!=null){ if ( cl_CHANIS0_5NormalHeightF.getCellType()==0) {CHANIS0_5NormalHeightF = "" + (int) cl_CHANIS0_5NormalHeightF.getNumericCellValue();} else if (cl_CHANIS0_5NormalHeightF.getCellType()==1) {CHANIS0_5NormalHeightF=cl_CHANIS0_5NormalHeightF.getStringCellValue();}  else if (cl_CHANIS0_5NormalHeightF.getCellType()== 2) {CHANIS0_5NormalHeightF=cl_CHANIS0_5NormalHeightF.getRawValue();  } else { CHANIS0_5NormalHeightF = "0";}}
XSSFCell cl_CHANIS0_5NormalHeightM = worksheet.getRow(69).getCell((short) 15); if ( cl_CHANIS0_5NormalHeightM!=null){ if ( cl_CHANIS0_5NormalHeightM.getCellType()==0) {CHANIS0_5NormalHeightM = "" + (int) cl_CHANIS0_5NormalHeightM.getNumericCellValue();} else if (cl_CHANIS0_5NormalHeightM.getCellType()==1) {CHANIS0_5NormalHeightM=cl_CHANIS0_5NormalHeightM.getStringCellValue();}  else if (cl_CHANIS0_5NormalHeightM.getCellType()== 2) {CHANIS0_5NormalHeightM=cl_CHANIS0_5NormalHeightM.getRawValue();  } else { CHANIS0_5NormalHeightM = "0";}}
XSSFCell cl_CHANIS0_5NormalHeightT = worksheet.getRow(69).getCell((short) 16); if ( cl_CHANIS0_5NormalHeightT!=null){ if ( cl_CHANIS0_5NormalHeightT.getCellType()==0) {CHANIS0_5NormalHeightT = "" + (int) cl_CHANIS0_5NormalHeightT.getNumericCellValue();} else if (cl_CHANIS0_5NormalHeightT.getCellType()==1) {CHANIS0_5NormalHeightT=cl_CHANIS0_5NormalHeightT.getStringCellValue();}  else if (cl_CHANIS0_5NormalHeightT.getCellType()== 2) {CHANIS0_5NormalHeightT=cl_CHANIS0_5NormalHeightT.getRawValue();  } else { CHANIS0_5NormalHeightT = "0";}}
XSSFCell cl_CHANIS0_5StuntedF = worksheet.getRow(70).getCell((short) 14); if ( cl_CHANIS0_5StuntedF!=null){ if ( cl_CHANIS0_5StuntedF.getCellType()==0) {CHANIS0_5StuntedF = "" + (int) cl_CHANIS0_5StuntedF.getNumericCellValue();} else if (cl_CHANIS0_5StuntedF.getCellType()==1) {CHANIS0_5StuntedF=cl_CHANIS0_5StuntedF.getStringCellValue();}  else if (cl_CHANIS0_5StuntedF.getCellType()== 2) {CHANIS0_5StuntedF=cl_CHANIS0_5StuntedF.getRawValue();  } else { CHANIS0_5StuntedF = "0";}}
XSSFCell cl_CHANIS0_5StuntedM = worksheet.getRow(70).getCell((short) 15); if ( cl_CHANIS0_5StuntedM!=null){ if ( cl_CHANIS0_5StuntedM.getCellType()==0) {CHANIS0_5StuntedM = "" + (int) cl_CHANIS0_5StuntedM.getNumericCellValue();} else if (cl_CHANIS0_5StuntedM.getCellType()==1) {CHANIS0_5StuntedM=cl_CHANIS0_5StuntedM.getStringCellValue();}  else if (cl_CHANIS0_5StuntedM.getCellType()== 2) {CHANIS0_5StuntedM=cl_CHANIS0_5StuntedM.getRawValue();  } else { CHANIS0_5StuntedM = "0";}}
XSSFCell cl_CHANIS0_5StuntedT = worksheet.getRow(70).getCell((short) 16); if ( cl_CHANIS0_5StuntedT!=null){ if ( cl_CHANIS0_5StuntedT.getCellType()==0) {CHANIS0_5StuntedT = "" + (int) cl_CHANIS0_5StuntedT.getNumericCellValue();} else if (cl_CHANIS0_5StuntedT.getCellType()==1) {CHANIS0_5StuntedT=cl_CHANIS0_5StuntedT.getStringCellValue();}  else if (cl_CHANIS0_5StuntedT.getCellType()== 2) {CHANIS0_5StuntedT=cl_CHANIS0_5StuntedT.getRawValue();  } else { CHANIS0_5StuntedT = "0";}}
XSSFCell cl_CHANIS0_5sevStuntedF = worksheet.getRow(71).getCell((short) 14); if ( cl_CHANIS0_5sevStuntedF!=null){ if ( cl_CHANIS0_5sevStuntedF.getCellType()==0) {CHANIS0_5sevStuntedF = "" + (int) cl_CHANIS0_5sevStuntedF.getNumericCellValue();} else if (cl_CHANIS0_5sevStuntedF.getCellType()==1) {CHANIS0_5sevStuntedF=cl_CHANIS0_5sevStuntedF.getStringCellValue();}  else if (cl_CHANIS0_5sevStuntedF.getCellType()== 2) {CHANIS0_5sevStuntedF=cl_CHANIS0_5sevStuntedF.getRawValue();  } else { CHANIS0_5sevStuntedF = "0";}}
XSSFCell cl_CHANIS0_5sevStuntedM = worksheet.getRow(71).getCell((short) 15); if ( cl_CHANIS0_5sevStuntedM!=null){ if ( cl_CHANIS0_5sevStuntedM.getCellType()==0) {CHANIS0_5sevStuntedM = "" + (int) cl_CHANIS0_5sevStuntedM.getNumericCellValue();} else if (cl_CHANIS0_5sevStuntedM.getCellType()==1) {CHANIS0_5sevStuntedM=cl_CHANIS0_5sevStuntedM.getStringCellValue();}  else if (cl_CHANIS0_5sevStuntedM.getCellType()== 2) {CHANIS0_5sevStuntedM=cl_CHANIS0_5sevStuntedM.getRawValue();  } else { CHANIS0_5sevStuntedM = "0";}}
XSSFCell cl_CHANIS0_5sevStuntedT = worksheet.getRow(71).getCell((short) 16); if ( cl_CHANIS0_5sevStuntedT!=null){ if ( cl_CHANIS0_5sevStuntedT.getCellType()==0) {CHANIS0_5sevStuntedT = "" + (int) cl_CHANIS0_5sevStuntedT.getNumericCellValue();} else if (cl_CHANIS0_5sevStuntedT.getCellType()==1) {CHANIS0_5sevStuntedT=cl_CHANIS0_5sevStuntedT.getStringCellValue();}  else if (cl_CHANIS0_5sevStuntedT.getCellType()== 2) {CHANIS0_5sevStuntedT=cl_CHANIS0_5sevStuntedT.getRawValue();  } else { CHANIS0_5sevStuntedT = "0";}}
XSSFCell cl_CHANIS0_5TMeasF = worksheet.getRow(72).getCell((short) 14); if ( cl_CHANIS0_5TMeasF!=null){ if ( cl_CHANIS0_5TMeasF.getCellType()==0) {CHANIS0_5TMeasF = "" + (int) cl_CHANIS0_5TMeasF.getNumericCellValue();} else if (cl_CHANIS0_5TMeasF.getCellType()==1) {CHANIS0_5TMeasF=cl_CHANIS0_5TMeasF.getStringCellValue();}  else if (cl_CHANIS0_5TMeasF.getCellType()== 2) {CHANIS0_5TMeasF=cl_CHANIS0_5TMeasF.getRawValue();  } else { CHANIS0_5TMeasF = "0";}}
XSSFCell cl_CHANIS0_5TMeasM = worksheet.getRow(72).getCell((short) 15); if ( cl_CHANIS0_5TMeasM!=null){ if ( cl_CHANIS0_5TMeasM.getCellType()==0) {CHANIS0_5TMeasM = "" + (int) cl_CHANIS0_5TMeasM.getNumericCellValue();} else if (cl_CHANIS0_5TMeasM.getCellType()==1) {CHANIS0_5TMeasM=cl_CHANIS0_5TMeasM.getStringCellValue();}  else if (cl_CHANIS0_5TMeasM.getCellType()== 2) {CHANIS0_5TMeasM=cl_CHANIS0_5TMeasM.getRawValue();  } else { CHANIS0_5TMeasM = "0";}}
XSSFCell cl_CHANIS0_5TMeas = worksheet.getRow(72).getCell((short) 16); if ( cl_CHANIS0_5TMeas!=null){ if ( cl_CHANIS0_5TMeas.getCellType()==0) {CHANIS0_5TMeas = "" + (int) cl_CHANIS0_5TMeas.getNumericCellValue();} else if (cl_CHANIS0_5TMeas.getCellType()==1) {CHANIS0_5TMeas=cl_CHANIS0_5TMeas.getStringCellValue();}  else if (cl_CHANIS0_5TMeas.getCellType()== 2) {CHANIS0_5TMeas=cl_CHANIS0_5TMeas.getRawValue();  } else { CHANIS0_5TMeas = "0";}}
XSSFCell cl_CHANIS6_23NormalHeightF = worksheet.getRow(73).getCell((short) 14); if ( cl_CHANIS6_23NormalHeightF!=null){ if ( cl_CHANIS6_23NormalHeightF.getCellType()==0) {CHANIS6_23NormalHeightF = "" + (int) cl_CHANIS6_23NormalHeightF.getNumericCellValue();} else if (cl_CHANIS6_23NormalHeightF.getCellType()==1) {CHANIS6_23NormalHeightF=cl_CHANIS6_23NormalHeightF.getStringCellValue();}  else if (cl_CHANIS6_23NormalHeightF.getCellType()== 2) {CHANIS6_23NormalHeightF=cl_CHANIS6_23NormalHeightF.getRawValue();  } else { CHANIS6_23NormalHeightF = "0";}}
XSSFCell cl_CHANIS6_23NormalHeightM = worksheet.getRow(73).getCell((short) 15); if ( cl_CHANIS6_23NormalHeightM!=null){ if ( cl_CHANIS6_23NormalHeightM.getCellType()==0) {CHANIS6_23NormalHeightM = "" + (int) cl_CHANIS6_23NormalHeightM.getNumericCellValue();} else if (cl_CHANIS6_23NormalHeightM.getCellType()==1) {CHANIS6_23NormalHeightM=cl_CHANIS6_23NormalHeightM.getStringCellValue();}  else if (cl_CHANIS6_23NormalHeightM.getCellType()== 2) {CHANIS6_23NormalHeightM=cl_CHANIS6_23NormalHeightM.getRawValue();  } else { CHANIS6_23NormalHeightM = "0";}}
XSSFCell cl_CHANIS6_23NormalHeightT = worksheet.getRow(73).getCell((short) 16); if ( cl_CHANIS6_23NormalHeightT!=null){ if ( cl_CHANIS6_23NormalHeightT.getCellType()==0) {CHANIS6_23NormalHeightT = "" + (int) cl_CHANIS6_23NormalHeightT.getNumericCellValue();} else if (cl_CHANIS6_23NormalHeightT.getCellType()==1) {CHANIS6_23NormalHeightT=cl_CHANIS6_23NormalHeightT.getStringCellValue();}  else if (cl_CHANIS6_23NormalHeightT.getCellType()== 2) {CHANIS6_23NormalHeightT=cl_CHANIS6_23NormalHeightT.getRawValue();  } else { CHANIS6_23NormalHeightT = "0";}}
XSSFCell cl_CHANIS6_23StuntedF = worksheet.getRow(74).getCell((short) 14); if ( cl_CHANIS6_23StuntedF!=null){ if ( cl_CHANIS6_23StuntedF.getCellType()==0) {CHANIS6_23StuntedF = "" + (int) cl_CHANIS6_23StuntedF.getNumericCellValue();} else if (cl_CHANIS6_23StuntedF.getCellType()==1) {CHANIS6_23StuntedF=cl_CHANIS6_23StuntedF.getStringCellValue();}  else if (cl_CHANIS6_23StuntedF.getCellType()== 2) {CHANIS6_23StuntedF=cl_CHANIS6_23StuntedF.getRawValue();  } else { CHANIS6_23StuntedF = "0";}}
XSSFCell cl_CHANIS6_23StuntedM = worksheet.getRow(74).getCell((short) 15); if ( cl_CHANIS6_23StuntedM!=null){ if ( cl_CHANIS6_23StuntedM.getCellType()==0) {CHANIS6_23StuntedM = "" + (int) cl_CHANIS6_23StuntedM.getNumericCellValue();} else if (cl_CHANIS6_23StuntedM.getCellType()==1) {CHANIS6_23StuntedM=cl_CHANIS6_23StuntedM.getStringCellValue();}  else if (cl_CHANIS6_23StuntedM.getCellType()== 2) {CHANIS6_23StuntedM=cl_CHANIS6_23StuntedM.getRawValue();  } else { CHANIS6_23StuntedM = "0";}}
XSSFCell cl_CHANIS6_23StuntedT = worksheet.getRow(74).getCell((short) 16); if ( cl_CHANIS6_23StuntedT!=null){ if ( cl_CHANIS6_23StuntedT.getCellType()==0) {CHANIS6_23StuntedT = "" + (int) cl_CHANIS6_23StuntedT.getNumericCellValue();} else if (cl_CHANIS6_23StuntedT.getCellType()==1) {CHANIS6_23StuntedT=cl_CHANIS6_23StuntedT.getStringCellValue();}  else if (cl_CHANIS6_23StuntedT.getCellType()== 2) {CHANIS6_23StuntedT=cl_CHANIS6_23StuntedT.getRawValue();  } else { CHANIS6_23StuntedT = "0";}}
XSSFCell cl_CHANIS6_23sevStuntedF = worksheet.getRow(75).getCell((short) 14); if ( cl_CHANIS6_23sevStuntedF!=null){ if ( cl_CHANIS6_23sevStuntedF.getCellType()==0) {CHANIS6_23sevStuntedF = "" + (int) cl_CHANIS6_23sevStuntedF.getNumericCellValue();} else if (cl_CHANIS6_23sevStuntedF.getCellType()==1) {CHANIS6_23sevStuntedF=cl_CHANIS6_23sevStuntedF.getStringCellValue();}  else if (cl_CHANIS6_23sevStuntedF.getCellType()== 2) {CHANIS6_23sevStuntedF=cl_CHANIS6_23sevStuntedF.getRawValue();  } else { CHANIS6_23sevStuntedF = "0";}}
XSSFCell cl_CHANIS6_23sevStuntedM = worksheet.getRow(75).getCell((short) 15); if ( cl_CHANIS6_23sevStuntedM!=null){ if ( cl_CHANIS6_23sevStuntedM.getCellType()==0) {CHANIS6_23sevStuntedM = "" + (int) cl_CHANIS6_23sevStuntedM.getNumericCellValue();} else if (cl_CHANIS6_23sevStuntedM.getCellType()==1) {CHANIS6_23sevStuntedM=cl_CHANIS6_23sevStuntedM.getStringCellValue();}  else if (cl_CHANIS6_23sevStuntedM.getCellType()== 2) {CHANIS6_23sevStuntedM=cl_CHANIS6_23sevStuntedM.getRawValue();  } else { CHANIS6_23sevStuntedM = "0";}}
XSSFCell cl_CHANIS6_23sevStuntedT = worksheet.getRow(75).getCell((short) 16); if ( cl_CHANIS6_23sevStuntedT!=null){ if ( cl_CHANIS6_23sevStuntedT.getCellType()==0) {CHANIS6_23sevStuntedT = "" + (int) cl_CHANIS6_23sevStuntedT.getNumericCellValue();} else if (cl_CHANIS6_23sevStuntedT.getCellType()==1) {CHANIS6_23sevStuntedT=cl_CHANIS6_23sevStuntedT.getStringCellValue();}  else if (cl_CHANIS6_23sevStuntedT.getCellType()== 2) {CHANIS6_23sevStuntedT=cl_CHANIS6_23sevStuntedT.getRawValue();  } else { CHANIS6_23sevStuntedT = "0";}}
XSSFCell cl_CHANIS6_23TMeasF = worksheet.getRow(76).getCell((short) 14); if ( cl_CHANIS6_23TMeasF!=null){ if ( cl_CHANIS6_23TMeasF.getCellType()==0) {CHANIS6_23TMeasF = "" + (int) cl_CHANIS6_23TMeasF.getNumericCellValue();} else if (cl_CHANIS6_23TMeasF.getCellType()==1) {CHANIS6_23TMeasF=cl_CHANIS6_23TMeasF.getStringCellValue();}  else if (cl_CHANIS6_23TMeasF.getCellType()== 2) {CHANIS6_23TMeasF=cl_CHANIS6_23TMeasF.getRawValue();  } else { CHANIS6_23TMeasF = "0";}}
XSSFCell cl_CHANIS6_23TMeasM = worksheet.getRow(76).getCell((short) 15); if ( cl_CHANIS6_23TMeasM!=null){ if ( cl_CHANIS6_23TMeasM.getCellType()==0) {CHANIS6_23TMeasM = "" + (int) cl_CHANIS6_23TMeasM.getNumericCellValue();} else if (cl_CHANIS6_23TMeasM.getCellType()==1) {CHANIS6_23TMeasM=cl_CHANIS6_23TMeasM.getStringCellValue();}  else if (cl_CHANIS6_23TMeasM.getCellType()== 2) {CHANIS6_23TMeasM=cl_CHANIS6_23TMeasM.getRawValue();  } else { CHANIS6_23TMeasM = "0";}}
XSSFCell cl_CHANIS6_23TMeas = worksheet.getRow(76).getCell((short) 16); if ( cl_CHANIS6_23TMeas!=null){ if ( cl_CHANIS6_23TMeas.getCellType()==0) {CHANIS6_23TMeas = "" + (int) cl_CHANIS6_23TMeas.getNumericCellValue();} else if (cl_CHANIS6_23TMeas.getCellType()==1) {CHANIS6_23TMeas=cl_CHANIS6_23TMeas.getStringCellValue();}  else if (cl_CHANIS6_23TMeas.getCellType()== 2) {CHANIS6_23TMeas=cl_CHANIS6_23TMeas.getRawValue();  } else { CHANIS6_23TMeas = "0";}}
XSSFCell cl_CHANIS24_59NormalHeightF = worksheet.getRow(77).getCell((short) 14); if ( cl_CHANIS24_59NormalHeightF!=null){ if ( cl_CHANIS24_59NormalHeightF.getCellType()==0) {CHANIS24_59NormalHeightF = "" + (int) cl_CHANIS24_59NormalHeightF.getNumericCellValue();} else if (cl_CHANIS24_59NormalHeightF.getCellType()==1) {CHANIS24_59NormalHeightF=cl_CHANIS24_59NormalHeightF.getStringCellValue();}  else if (cl_CHANIS24_59NormalHeightF.getCellType()== 2) {CHANIS24_59NormalHeightF=cl_CHANIS24_59NormalHeightF.getRawValue();  } else { CHANIS24_59NormalHeightF = "0";}}
XSSFCell cl_CHANIS24_59NormalHeightM = worksheet.getRow(77).getCell((short) 15); if ( cl_CHANIS24_59NormalHeightM!=null){ if ( cl_CHANIS24_59NormalHeightM.getCellType()==0) {CHANIS24_59NormalHeightM = "" + (int) cl_CHANIS24_59NormalHeightM.getNumericCellValue();} else if (cl_CHANIS24_59NormalHeightM.getCellType()==1) {CHANIS24_59NormalHeightM=cl_CHANIS24_59NormalHeightM.getStringCellValue();}  else if (cl_CHANIS24_59NormalHeightM.getCellType()== 2) {CHANIS24_59NormalHeightM=cl_CHANIS24_59NormalHeightM.getRawValue();  } else { CHANIS24_59NormalHeightM = "0";}}
XSSFCell cl_CHANIS24_59NormalHeightT = worksheet.getRow(77).getCell((short) 16); if ( cl_CHANIS24_59NormalHeightT!=null){ if ( cl_CHANIS24_59NormalHeightT.getCellType()==0) {CHANIS24_59NormalHeightT = "" + (int) cl_CHANIS24_59NormalHeightT.getNumericCellValue();} else if (cl_CHANIS24_59NormalHeightT.getCellType()==1) {CHANIS24_59NormalHeightT=cl_CHANIS24_59NormalHeightT.getStringCellValue();}  else if (cl_CHANIS24_59NormalHeightT.getCellType()== 2) {CHANIS24_59NormalHeightT=cl_CHANIS24_59NormalHeightT.getRawValue();  } else { CHANIS24_59NormalHeightT = "0";}}
XSSFCell cl_CHANIS24_59StuntedF = worksheet.getRow(78).getCell((short) 14); if ( cl_CHANIS24_59StuntedF!=null){ if ( cl_CHANIS24_59StuntedF.getCellType()==0) {CHANIS24_59StuntedF = "" + (int) cl_CHANIS24_59StuntedF.getNumericCellValue();} else if (cl_CHANIS24_59StuntedF.getCellType()==1) {CHANIS24_59StuntedF=cl_CHANIS24_59StuntedF.getStringCellValue();}  else if (cl_CHANIS24_59StuntedF.getCellType()== 2) {CHANIS24_59StuntedF=cl_CHANIS24_59StuntedF.getRawValue();  } else { CHANIS24_59StuntedF = "0";}}
XSSFCell cl_CHANIS24_59StuntedM = worksheet.getRow(78).getCell((short) 15); if ( cl_CHANIS24_59StuntedM!=null){ if ( cl_CHANIS24_59StuntedM.getCellType()==0) {CHANIS24_59StuntedM = "" + (int) cl_CHANIS24_59StuntedM.getNumericCellValue();} else if (cl_CHANIS24_59StuntedM.getCellType()==1) {CHANIS24_59StuntedM=cl_CHANIS24_59StuntedM.getStringCellValue();}  else if (cl_CHANIS24_59StuntedM.getCellType()== 2) {CHANIS24_59StuntedM=cl_CHANIS24_59StuntedM.getRawValue();  } else { CHANIS24_59StuntedM = "0";}}
XSSFCell cl_CHANIS24_59StuntedT = worksheet.getRow(78).getCell((short) 16); if ( cl_CHANIS24_59StuntedT!=null){ if ( cl_CHANIS24_59StuntedT.getCellType()==0) {CHANIS24_59StuntedT = "" + (int) cl_CHANIS24_59StuntedT.getNumericCellValue();} else if (cl_CHANIS24_59StuntedT.getCellType()==1) {CHANIS24_59StuntedT=cl_CHANIS24_59StuntedT.getStringCellValue();}  else if (cl_CHANIS24_59StuntedT.getCellType()== 2) {CHANIS24_59StuntedT=cl_CHANIS24_59StuntedT.getRawValue();  } else { CHANIS24_59StuntedT = "0";}}
XSSFCell cl_CHANIS24_59sevStuntedF = worksheet.getRow(79).getCell((short) 14); if ( cl_CHANIS24_59sevStuntedF!=null){ if ( cl_CHANIS24_59sevStuntedF.getCellType()==0) {CHANIS24_59sevStuntedF = "" + (int) cl_CHANIS24_59sevStuntedF.getNumericCellValue();} else if (cl_CHANIS24_59sevStuntedF.getCellType()==1) {CHANIS24_59sevStuntedF=cl_CHANIS24_59sevStuntedF.getStringCellValue();}  else if (cl_CHANIS24_59sevStuntedF.getCellType()== 2) {CHANIS24_59sevStuntedF=cl_CHANIS24_59sevStuntedF.getRawValue();  } else { CHANIS24_59sevStuntedF = "0";}}
XSSFCell cl_CHANIS24_59sevStuntedM = worksheet.getRow(79).getCell((short) 15); if ( cl_CHANIS24_59sevStuntedM!=null){ if ( cl_CHANIS24_59sevStuntedM.getCellType()==0) {CHANIS24_59sevStuntedM = "" + (int) cl_CHANIS24_59sevStuntedM.getNumericCellValue();} else if (cl_CHANIS24_59sevStuntedM.getCellType()==1) {CHANIS24_59sevStuntedM=cl_CHANIS24_59sevStuntedM.getStringCellValue();}  else if (cl_CHANIS24_59sevStuntedM.getCellType()== 2) {CHANIS24_59sevStuntedM=cl_CHANIS24_59sevStuntedM.getRawValue();  } else { CHANIS24_59sevStuntedM = "0";}}
XSSFCell cl_CHANIS24_59sevStuntedT = worksheet.getRow(79).getCell((short) 16); if ( cl_CHANIS24_59sevStuntedT!=null){ if ( cl_CHANIS24_59sevStuntedT.getCellType()==0) {CHANIS24_59sevStuntedT = "" + (int) cl_CHANIS24_59sevStuntedT.getNumericCellValue();} else if (cl_CHANIS24_59sevStuntedT.getCellType()==1) {CHANIS24_59sevStuntedT=cl_CHANIS24_59sevStuntedT.getStringCellValue();}  else if (cl_CHANIS24_59sevStuntedT.getCellType()== 2) {CHANIS24_59sevStuntedT=cl_CHANIS24_59sevStuntedT.getRawValue();  } else { CHANIS24_59sevStuntedT = "0";}}
XSSFCell cl_CHANIS24_59TMeasF = worksheet.getRow(80).getCell((short) 14); if ( cl_CHANIS24_59TMeasF!=null){ if ( cl_CHANIS24_59TMeasF.getCellType()==0) {CHANIS24_59TMeasF = "" + (int) cl_CHANIS24_59TMeasF.getNumericCellValue();} else if (cl_CHANIS24_59TMeasF.getCellType()==1) {CHANIS24_59TMeasF=cl_CHANIS24_59TMeasF.getStringCellValue();}  else if (cl_CHANIS24_59TMeasF.getCellType()== 2) {CHANIS24_59TMeasF=cl_CHANIS24_59TMeasF.getRawValue();  } else { CHANIS24_59TMeasF = "0";}}
XSSFCell cl_CHANIS24_59TMeasM = worksheet.getRow(80).getCell((short) 15); if ( cl_CHANIS24_59TMeasM!=null){ if ( cl_CHANIS24_59TMeasM.getCellType()==0) {CHANIS24_59TMeasM = "" + (int) cl_CHANIS24_59TMeasM.getNumericCellValue();} else if (cl_CHANIS24_59TMeasM.getCellType()==1) {CHANIS24_59TMeasM=cl_CHANIS24_59TMeasM.getStringCellValue();}  else if (cl_CHANIS24_59TMeasM.getCellType()== 2) {CHANIS24_59TMeasM=cl_CHANIS24_59TMeasM.getRawValue();  } else { CHANIS24_59TMeasM = "0";}}
XSSFCell cl_CHANIS24_59TMeas = worksheet.getRow(80).getCell((short) 16); if ( cl_CHANIS24_59TMeas!=null){ if ( cl_CHANIS24_59TMeas.getCellType()==0) {CHANIS24_59TMeas = "" + (int) cl_CHANIS24_59TMeas.getNumericCellValue();} else if (cl_CHANIS24_59TMeas.getCellType()==1) {CHANIS24_59TMeas=cl_CHANIS24_59TMeas.getStringCellValue();}  else if (cl_CHANIS24_59TMeas.getCellType()== 2) {CHANIS24_59TMeas=cl_CHANIS24_59TMeas.getRawValue();  } else { CHANIS24_59TMeas = "0";}}
XSSFCell cl_CHANIS0_59NewVisitsF = worksheet.getRow(82).getCell((short) 14); if ( cl_CHANIS0_59NewVisitsF!=null){ if ( cl_CHANIS0_59NewVisitsF.getCellType()==0) {CHANIS0_59NewVisitsF = "" + (int) cl_CHANIS0_59NewVisitsF.getNumericCellValue();} else if (cl_CHANIS0_59NewVisitsF.getCellType()==1) {CHANIS0_59NewVisitsF=cl_CHANIS0_59NewVisitsF.getStringCellValue();}  else if (cl_CHANIS0_59NewVisitsF.getCellType()== 2) {CHANIS0_59NewVisitsF=cl_CHANIS0_59NewVisitsF.getRawValue();  } else { CHANIS0_59NewVisitsF = "0";}}
XSSFCell cl_CHANIS0_59NewVisitsM = worksheet.getRow(82).getCell((short) 15); if ( cl_CHANIS0_59NewVisitsM!=null){ if ( cl_CHANIS0_59NewVisitsM.getCellType()==0) {CHANIS0_59NewVisitsM = "" + (int) cl_CHANIS0_59NewVisitsM.getNumericCellValue();} else if (cl_CHANIS0_59NewVisitsM.getCellType()==1) {CHANIS0_59NewVisitsM=cl_CHANIS0_59NewVisitsM.getStringCellValue();}  else if (cl_CHANIS0_59NewVisitsM.getCellType()== 2) {CHANIS0_59NewVisitsM=cl_CHANIS0_59NewVisitsM.getRawValue();  } else { CHANIS0_59NewVisitsM = "0";}}
XSSFCell cl_CHANIS0_59NewVisitsT = worksheet.getRow(82).getCell((short) 16); if ( cl_CHANIS0_59NewVisitsT!=null){ if ( cl_CHANIS0_59NewVisitsT.getCellType()==0) {CHANIS0_59NewVisitsT = "" + (int) cl_CHANIS0_59NewVisitsT.getNumericCellValue();} else if (cl_CHANIS0_59NewVisitsT.getCellType()==1) {CHANIS0_59NewVisitsT=cl_CHANIS0_59NewVisitsT.getStringCellValue();}  else if (cl_CHANIS0_59NewVisitsT.getCellType()== 2) {CHANIS0_59NewVisitsT=cl_CHANIS0_59NewVisitsT.getRawValue();  } else { CHANIS0_59NewVisitsT = "0";}}
XSSFCell cl_CHANIS0_59KwashiakorF = worksheet.getRow(83).getCell((short) 14); if ( cl_CHANIS0_59KwashiakorF!=null){ if ( cl_CHANIS0_59KwashiakorF.getCellType()==0) {CHANIS0_59KwashiakorF = "" + (int) cl_CHANIS0_59KwashiakorF.getNumericCellValue();} else if (cl_CHANIS0_59KwashiakorF.getCellType()==1) {CHANIS0_59KwashiakorF=cl_CHANIS0_59KwashiakorF.getStringCellValue();}  else if (cl_CHANIS0_59KwashiakorF.getCellType()== 2) {CHANIS0_59KwashiakorF=cl_CHANIS0_59KwashiakorF.getRawValue();  } else { CHANIS0_59KwashiakorF = "0";}}
XSSFCell cl_CHANIS0_59KwashiakorM = worksheet.getRow(83).getCell((short) 15); if ( cl_CHANIS0_59KwashiakorM!=null){ if ( cl_CHANIS0_59KwashiakorM.getCellType()==0) {CHANIS0_59KwashiakorM = "" + (int) cl_CHANIS0_59KwashiakorM.getNumericCellValue();} else if (cl_CHANIS0_59KwashiakorM.getCellType()==1) {CHANIS0_59KwashiakorM=cl_CHANIS0_59KwashiakorM.getStringCellValue();}  else if (cl_CHANIS0_59KwashiakorM.getCellType()== 2) {CHANIS0_59KwashiakorM=cl_CHANIS0_59KwashiakorM.getRawValue();  } else { CHANIS0_59KwashiakorM = "0";}}
XSSFCell cl_CHANIS0_59KwashiakorT = worksheet.getRow(83).getCell((short) 16); if ( cl_CHANIS0_59KwashiakorT!=null){ if ( cl_CHANIS0_59KwashiakorT.getCellType()==0) {CHANIS0_59KwashiakorT = "" + (int) cl_CHANIS0_59KwashiakorT.getNumericCellValue();} else if (cl_CHANIS0_59KwashiakorT.getCellType()==1) {CHANIS0_59KwashiakorT=cl_CHANIS0_59KwashiakorT.getStringCellValue();}  else if (cl_CHANIS0_59KwashiakorT.getCellType()== 2) {CHANIS0_59KwashiakorT=cl_CHANIS0_59KwashiakorT.getRawValue();  } else { CHANIS0_59KwashiakorT = "0";}}
XSSFCell cl_CHANIS0_59MarasmusF = worksheet.getRow(84).getCell((short) 14); if ( cl_CHANIS0_59MarasmusF!=null){ if ( cl_CHANIS0_59MarasmusF.getCellType()==0) {CHANIS0_59MarasmusF = "" + (int) cl_CHANIS0_59MarasmusF.getNumericCellValue();} else if (cl_CHANIS0_59MarasmusF.getCellType()==1) {CHANIS0_59MarasmusF=cl_CHANIS0_59MarasmusF.getStringCellValue();}  else if (cl_CHANIS0_59MarasmusF.getCellType()== 2) {CHANIS0_59MarasmusF=cl_CHANIS0_59MarasmusF.getRawValue();  } else { CHANIS0_59MarasmusF = "0";}}
XSSFCell cl_CHANIS0_59MarasmusM = worksheet.getRow(84).getCell((short) 15); if ( cl_CHANIS0_59MarasmusM!=null){ if ( cl_CHANIS0_59MarasmusM.getCellType()==0) {CHANIS0_59MarasmusM = "" + (int) cl_CHANIS0_59MarasmusM.getNumericCellValue();} else if (cl_CHANIS0_59MarasmusM.getCellType()==1) {CHANIS0_59MarasmusM=cl_CHANIS0_59MarasmusM.getStringCellValue();}  else if (cl_CHANIS0_59MarasmusM.getCellType()== 2) {CHANIS0_59MarasmusM=cl_CHANIS0_59MarasmusM.getRawValue();  } else { CHANIS0_59MarasmusM = "0";}}
XSSFCell cl_CHANIS0_59MarasmusT = worksheet.getRow(84).getCell((short) 16); if ( cl_CHANIS0_59MarasmusT!=null){ if ( cl_CHANIS0_59MarasmusT.getCellType()==0) {CHANIS0_59MarasmusT = "" + (int) cl_CHANIS0_59MarasmusT.getNumericCellValue();} else if (cl_CHANIS0_59MarasmusT.getCellType()==1) {CHANIS0_59MarasmusT=cl_CHANIS0_59MarasmusT.getStringCellValue();}  else if (cl_CHANIS0_59MarasmusT.getCellType()== 2) {CHANIS0_59MarasmusT=cl_CHANIS0_59MarasmusT.getRawValue();  } else { CHANIS0_59MarasmusT = "0";}}
XSSFCell cl_CHANIS0_59FalgrowthF = worksheet.getRow(85).getCell((short) 14); if ( cl_CHANIS0_59FalgrowthF!=null){ if ( cl_CHANIS0_59FalgrowthF.getCellType()==0) {CHANIS0_59FalgrowthF = "" + (int) cl_CHANIS0_59FalgrowthF.getNumericCellValue();} else if (cl_CHANIS0_59FalgrowthF.getCellType()==1) {CHANIS0_59FalgrowthF=cl_CHANIS0_59FalgrowthF.getStringCellValue();}  else if (cl_CHANIS0_59FalgrowthF.getCellType()== 2) {CHANIS0_59FalgrowthF=cl_CHANIS0_59FalgrowthF.getRawValue();  } else { CHANIS0_59FalgrowthF = "0";}}
XSSFCell cl_CHANIS0_59FalgrowthM = worksheet.getRow(85).getCell((short) 15); if ( cl_CHANIS0_59FalgrowthM!=null){ if ( cl_CHANIS0_59FalgrowthM.getCellType()==0) {CHANIS0_59FalgrowthM = "" + (int) cl_CHANIS0_59FalgrowthM.getNumericCellValue();} else if (cl_CHANIS0_59FalgrowthM.getCellType()==1) {CHANIS0_59FalgrowthM=cl_CHANIS0_59FalgrowthM.getStringCellValue();}  else if (cl_CHANIS0_59FalgrowthM.getCellType()== 2) {CHANIS0_59FalgrowthM=cl_CHANIS0_59FalgrowthM.getRawValue();  } else { CHANIS0_59FalgrowthM = "0";}}
XSSFCell cl_CHANIS0_59FalgrowthT = worksheet.getRow(85).getCell((short) 16); if ( cl_CHANIS0_59FalgrowthT!=null){ if ( cl_CHANIS0_59FalgrowthT.getCellType()==0) {CHANIS0_59FalgrowthT = "" + (int) cl_CHANIS0_59FalgrowthT.getNumericCellValue();} else if (cl_CHANIS0_59FalgrowthT.getCellType()==1) {CHANIS0_59FalgrowthT=cl_CHANIS0_59FalgrowthT.getStringCellValue();}  else if (cl_CHANIS0_59FalgrowthT.getCellType()== 2) {CHANIS0_59FalgrowthT=cl_CHANIS0_59FalgrowthT.getRawValue();  } else { CHANIS0_59FalgrowthT = "0";}}
XSSFCell cl_CHANIS0_59F = worksheet.getRow(86).getCell((short) 14); if ( cl_CHANIS0_59F!=null){ if ( cl_CHANIS0_59F.getCellType()==0) {CHANIS0_59F = "" + (int) cl_CHANIS0_59F.getNumericCellValue();} else if (cl_CHANIS0_59F.getCellType()==1) {CHANIS0_59F=cl_CHANIS0_59F.getStringCellValue();}  else if (cl_CHANIS0_59F.getCellType()== 2) {CHANIS0_59F=cl_CHANIS0_59F.getRawValue();  } else { CHANIS0_59F = "0";}}
XSSFCell cl_CHANIS0_59M = worksheet.getRow(86).getCell((short) 15); if ( cl_CHANIS0_59M!=null){ if ( cl_CHANIS0_59M.getCellType()==0) {CHANIS0_59M = "" + (int) cl_CHANIS0_59M.getNumericCellValue();} else if (cl_CHANIS0_59M.getCellType()==1) {CHANIS0_59M=cl_CHANIS0_59M.getStringCellValue();}  else if (cl_CHANIS0_59M.getCellType()== 2) {CHANIS0_59M=cl_CHANIS0_59M.getRawValue();  } else { CHANIS0_59M = "0";}}
XSSFCell cl_CHANIS0_59T = worksheet.getRow(86).getCell((short) 16); if ( cl_CHANIS0_59T!=null){ if ( cl_CHANIS0_59T.getCellType()==0) {CHANIS0_59T = "" + (int) cl_CHANIS0_59T.getNumericCellValue();} else if (cl_CHANIS0_59T.getCellType()==1) {CHANIS0_59T=cl_CHANIS0_59T.getStringCellValue();}  else if (cl_CHANIS0_59T.getCellType()== 2) {CHANIS0_59T=cl_CHANIS0_59T.getRawValue();  } else { CHANIS0_59T = "0";}}
XSSFCell cl_CHANIS0_5EXCLBreastF = worksheet.getRow(87).getCell((short) 14); if ( cl_CHANIS0_5EXCLBreastF!=null){ if ( cl_CHANIS0_5EXCLBreastF.getCellType()==0) {CHANIS0_5EXCLBreastF = "" + (int) cl_CHANIS0_5EXCLBreastF.getNumericCellValue();} else if (cl_CHANIS0_5EXCLBreastF.getCellType()==1) {CHANIS0_5EXCLBreastF=cl_CHANIS0_5EXCLBreastF.getStringCellValue();}  else if (cl_CHANIS0_5EXCLBreastF.getCellType()== 2) {CHANIS0_5EXCLBreastF=cl_CHANIS0_5EXCLBreastF.getRawValue();  } else { CHANIS0_5EXCLBreastF = "0";}}
XSSFCell cl_CHANIS0_5EXCLBreastM = worksheet.getRow(87).getCell((short) 15); if ( cl_CHANIS0_5EXCLBreastM!=null){ if ( cl_CHANIS0_5EXCLBreastM.getCellType()==0) {CHANIS0_5EXCLBreastM = "" + (int) cl_CHANIS0_5EXCLBreastM.getNumericCellValue();} else if (cl_CHANIS0_5EXCLBreastM.getCellType()==1) {CHANIS0_5EXCLBreastM=cl_CHANIS0_5EXCLBreastM.getStringCellValue();}  else if (cl_CHANIS0_5EXCLBreastM.getCellType()== 2) {CHANIS0_5EXCLBreastM=cl_CHANIS0_5EXCLBreastM.getRawValue();  } else { CHANIS0_5EXCLBreastM = "0";}}
XSSFCell cl_CHANIS0_5EXCLBreastT = worksheet.getRow(87).getCell((short) 16); if ( cl_CHANIS0_5EXCLBreastT!=null){ if ( cl_CHANIS0_5EXCLBreastT.getCellType()==0) {CHANIS0_5EXCLBreastT = "" + (int) cl_CHANIS0_5EXCLBreastT.getNumericCellValue();} else if (cl_CHANIS0_5EXCLBreastT.getCellType()==1) {CHANIS0_5EXCLBreastT=cl_CHANIS0_5EXCLBreastT.getStringCellValue();}  else if (cl_CHANIS0_5EXCLBreastT.getCellType()== 2) {CHANIS0_5EXCLBreastT=cl_CHANIS0_5EXCLBreastT.getRawValue();  } else { CHANIS0_5EXCLBreastT = "0";}}
XSSFCell cl_CHANIS12_59DewormedF = worksheet.getRow(88).getCell((short) 14); if ( cl_CHANIS12_59DewormedF!=null){ if ( cl_CHANIS12_59DewormedF.getCellType()==0) {CHANIS12_59DewormedF = "" + (int) cl_CHANIS12_59DewormedF.getNumericCellValue();} else if (cl_CHANIS12_59DewormedF.getCellType()==1) {CHANIS12_59DewormedF=cl_CHANIS12_59DewormedF.getStringCellValue();}  else if (cl_CHANIS12_59DewormedF.getCellType()== 2) {CHANIS12_59DewormedF=cl_CHANIS12_59DewormedF.getRawValue();  } else { CHANIS12_59DewormedF = "0";}}
XSSFCell cl_CHANIS12_59DewormedM = worksheet.getRow(88).getCell((short) 15); if ( cl_CHANIS12_59DewormedM!=null){ if ( cl_CHANIS12_59DewormedM.getCellType()==0) {CHANIS12_59DewormedM = "" + (int) cl_CHANIS12_59DewormedM.getNumericCellValue();} else if (cl_CHANIS12_59DewormedM.getCellType()==1) {CHANIS12_59DewormedM=cl_CHANIS12_59DewormedM.getStringCellValue();}  else if (cl_CHANIS12_59DewormedM.getCellType()== 2) {CHANIS12_59DewormedM=cl_CHANIS12_59DewormedM.getRawValue();  } else { CHANIS12_59DewormedM = "0";}}
XSSFCell cl_CHANIS12_59DewormedT = worksheet.getRow(88).getCell((short) 16); if ( cl_CHANIS12_59DewormedT!=null){ if ( cl_CHANIS12_59DewormedT.getCellType()==0) {CHANIS12_59DewormedT = "" + (int) cl_CHANIS12_59DewormedT.getNumericCellValue();} else if (cl_CHANIS12_59DewormedT.getCellType()==1) {CHANIS12_59DewormedT=cl_CHANIS12_59DewormedT.getStringCellValue();}  else if (cl_CHANIS12_59DewormedT.getCellType()== 2) {CHANIS12_59DewormedT=cl_CHANIS12_59DewormedT.getRawValue();  } else { CHANIS12_59DewormedT = "0";}}
XSSFCell cl_CHANIS6_23MNPsF = worksheet.getRow(89).getCell((short) 14); if ( cl_CHANIS6_23MNPsF!=null){ if ( cl_CHANIS6_23MNPsF.getCellType()==0) {CHANIS6_23MNPsF = "" + (int) cl_CHANIS6_23MNPsF.getNumericCellValue();} else if (cl_CHANIS6_23MNPsF.getCellType()==1) {CHANIS6_23MNPsF=cl_CHANIS6_23MNPsF.getStringCellValue();}  else if (cl_CHANIS6_23MNPsF.getCellType()== 2) {CHANIS6_23MNPsF=cl_CHANIS6_23MNPsF.getRawValue();  } else { CHANIS6_23MNPsF = "0";}}
XSSFCell cl_CHANIS6_23MNPsM = worksheet.getRow(89).getCell((short) 15); if ( cl_CHANIS6_23MNPsM!=null){ if ( cl_CHANIS6_23MNPsM.getCellType()==0) {CHANIS6_23MNPsM = "" + (int) cl_CHANIS6_23MNPsM.getNumericCellValue();} else if (cl_CHANIS6_23MNPsM.getCellType()==1) {CHANIS6_23MNPsM=cl_CHANIS6_23MNPsM.getStringCellValue();}  else if (cl_CHANIS6_23MNPsM.getCellType()== 2) {CHANIS6_23MNPsM=cl_CHANIS6_23MNPsM.getRawValue();  } else { CHANIS6_23MNPsM = "0";}}
XSSFCell cl_CHANIS6_23MNPsT = worksheet.getRow(89).getCell((short) 16); if ( cl_CHANIS6_23MNPsT!=null){ if ( cl_CHANIS6_23MNPsT.getCellType()==0) {CHANIS6_23MNPsT = "" + (int) cl_CHANIS6_23MNPsT.getNumericCellValue();} else if (cl_CHANIS6_23MNPsT.getCellType()==1) {CHANIS6_23MNPsT=cl_CHANIS6_23MNPsT.getStringCellValue();}  else if (cl_CHANIS6_23MNPsT.getCellType()== 2) {CHANIS6_23MNPsT=cl_CHANIS6_23MNPsT.getRawValue();  } else { CHANIS6_23MNPsT = "0";}}
XSSFCell cl_CHANIS0_59DisabilityF = worksheet.getRow(90).getCell((short) 14); if ( cl_CHANIS0_59DisabilityF!=null){ if ( cl_CHANIS0_59DisabilityF.getCellType()==0) {CHANIS0_59DisabilityF = "" + (int) cl_CHANIS0_59DisabilityF.getNumericCellValue();} else if (cl_CHANIS0_59DisabilityF.getCellType()==1) {CHANIS0_59DisabilityF=cl_CHANIS0_59DisabilityF.getStringCellValue();}  else if (cl_CHANIS0_59DisabilityF.getCellType()== 2) {CHANIS0_59DisabilityF=cl_CHANIS0_59DisabilityF.getRawValue();  } else { CHANIS0_59DisabilityF = "0";}}
XSSFCell cl_CHANIS0_59DisabilityM = worksheet.getRow(90).getCell((short) 15); if ( cl_CHANIS0_59DisabilityM!=null){ if ( cl_CHANIS0_59DisabilityM.getCellType()==0) {CHANIS0_59DisabilityM = "" + (int) cl_CHANIS0_59DisabilityM.getNumericCellValue();} else if (cl_CHANIS0_59DisabilityM.getCellType()==1) {CHANIS0_59DisabilityM=cl_CHANIS0_59DisabilityM.getStringCellValue();}  else if (cl_CHANIS0_59DisabilityM.getCellType()== 2) {CHANIS0_59DisabilityM=cl_CHANIS0_59DisabilityM.getRawValue();  } else { CHANIS0_59DisabilityM = "0";}}
XSSFCell cl_CHANIS0_59DisabilityT = worksheet.getRow(90).getCell((short) 16); if ( cl_CHANIS0_59DisabilityT!=null){ if ( cl_CHANIS0_59DisabilityT.getCellType()==0) {CHANIS0_59DisabilityT = "" + (int) cl_CHANIS0_59DisabilityT.getNumericCellValue();} else if (cl_CHANIS0_59DisabilityT.getCellType()==1) {CHANIS0_59DisabilityT=cl_CHANIS0_59DisabilityT.getStringCellValue();}  else if (cl_CHANIS0_59DisabilityT.getCellType()== 2) {CHANIS0_59DisabilityT=cl_CHANIS0_59DisabilityT.getRawValue();  } else { CHANIS0_59DisabilityT = "0";}}
XSSFCell cl_CCSVVH24 = worksheet.getRow(54).getCell((short) 7); if ( cl_CCSVVH24!=null){ if ( cl_CCSVVH24.getCellType()==0) {CCSVVH24 = "" + (int) cl_CCSVVH24.getNumericCellValue();} else if (cl_CCSVVH24.getCellType()==1) {CCSVVH24=cl_CCSVVH24.getStringCellValue();}  else if (cl_CCSVVH24.getCellType()== 2) {CCSVVH24=cl_CCSVVH24.getRawValue();  } else { CCSVVH24 = "0";}}
XSSFCell cl_CCSVVH25_49 = worksheet.getRow(54).getCell((short) 8); if ( cl_CCSVVH25_49!=null){ if ( cl_CCSVVH25_49.getCellType()==0) {CCSVVH25_49 = "" + (int) cl_CCSVVH25_49.getNumericCellValue();} else if (cl_CCSVVH25_49.getCellType()==1) {CCSVVH25_49=cl_CCSVVH25_49.getStringCellValue();}  else if (cl_CCSVVH25_49.getCellType()== 2) {CCSVVH25_49=cl_CCSVVH25_49.getRawValue();  } else { CCSVVH25_49 = "0";}}
XSSFCell cl_CCSVVH50 = worksheet.getRow(54).getCell((short) 9); if ( cl_CCSVVH50!=null){ if ( cl_CCSVVH50.getCellType()==0) {CCSVVH50 = "" + (int) cl_CCSVVH50.getNumericCellValue();} else if (cl_CCSVVH50.getCellType()==1) {CCSVVH50=cl_CCSVVH50.getStringCellValue();}  else if (cl_CCSVVH50.getCellType()== 2) {CCSVVH50=cl_CCSVVH50.getRawValue();  } else { CCSVVH50 = "0";}}
XSSFCell cl_CCSPAPSMEAR24 = worksheet.getRow(55).getCell((short) 7); if ( cl_CCSPAPSMEAR24!=null){ if ( cl_CCSPAPSMEAR24.getCellType()==0) {CCSPAPSMEAR24 = "" + (int) cl_CCSPAPSMEAR24.getNumericCellValue();} else if (cl_CCSPAPSMEAR24.getCellType()==1) {CCSPAPSMEAR24=cl_CCSPAPSMEAR24.getStringCellValue();}  else if (cl_CCSPAPSMEAR24.getCellType()== 2) {CCSPAPSMEAR24=cl_CCSPAPSMEAR24.getRawValue();  } else { CCSPAPSMEAR24 = "0";}}
XSSFCell cl_CCSPAPSMEAR25_49 = worksheet.getRow(55).getCell((short) 8); if ( cl_CCSPAPSMEAR25_49!=null){ if ( cl_CCSPAPSMEAR25_49.getCellType()==0) {CCSPAPSMEAR25_49 = "" + (int) cl_CCSPAPSMEAR25_49.getNumericCellValue();} else if (cl_CCSPAPSMEAR25_49.getCellType()==1) {CCSPAPSMEAR25_49=cl_CCSPAPSMEAR25_49.getStringCellValue();}  else if (cl_CCSPAPSMEAR25_49.getCellType()== 2) {CCSPAPSMEAR25_49=cl_CCSPAPSMEAR25_49.getRawValue();  } else { CCSPAPSMEAR25_49 = "0";}}
XSSFCell cl_CCSPAPSMEAR50 = worksheet.getRow(55).getCell((short) 9); if ( cl_CCSPAPSMEAR50!=null){ if ( cl_CCSPAPSMEAR50.getCellType()==0) {CCSPAPSMEAR50 = "" + (int) cl_CCSPAPSMEAR50.getNumericCellValue();} else if (cl_CCSPAPSMEAR50.getCellType()==1) {CCSPAPSMEAR50=cl_CCSPAPSMEAR50.getStringCellValue();}  else if (cl_CCSPAPSMEAR50.getCellType()== 2) {CCSPAPSMEAR50=cl_CCSPAPSMEAR50.getRawValue();  } else { CCSPAPSMEAR50 = "0";}}
XSSFCell cl_CCSHPV24 = worksheet.getRow(56).getCell((short) 7); if ( cl_CCSHPV24!=null){ if ( cl_CCSHPV24.getCellType()==0) {CCSHPV24 = "" + (int) cl_CCSHPV24.getNumericCellValue();} else if (cl_CCSHPV24.getCellType()==1) {CCSHPV24=cl_CCSHPV24.getStringCellValue();}  else if (cl_CCSHPV24.getCellType()== 2) {CCSHPV24=cl_CCSHPV24.getRawValue();  } else { CCSHPV24 = "0";}}
XSSFCell cl_CCSHPV25_49 = worksheet.getRow(56).getCell((short) 8); if ( cl_CCSHPV25_49!=null){ if ( cl_CCSHPV25_49.getCellType()==0) {CCSHPV25_49 = "" + (int) cl_CCSHPV25_49.getNumericCellValue();} else if (cl_CCSHPV25_49.getCellType()==1) {CCSHPV25_49=cl_CCSHPV25_49.getStringCellValue();}  else if (cl_CCSHPV25_49.getCellType()== 2) {CCSHPV25_49=cl_CCSHPV25_49.getRawValue();  } else { CCSHPV25_49 = "0";}}
XSSFCell cl_CCSHPV50 = worksheet.getRow(56).getCell((short) 9); if ( cl_CCSHPV50!=null){ if ( cl_CCSHPV50.getCellType()==0) {CCSHPV50 = "" + (int) cl_CCSHPV50.getNumericCellValue();} else if (cl_CCSHPV50.getCellType()==1) {CCSHPV50=cl_CCSHPV50.getStringCellValue();}  else if (cl_CCSHPV50.getCellType()== 2) {CCSHPV50=cl_CCSHPV50.getRawValue();  } else { CCSHPV50 = "0";}}
XSSFCell cl_CCSVIAVILIPOS24 = worksheet.getRow(57).getCell((short) 7); if ( cl_CCSVIAVILIPOS24!=null){ if ( cl_CCSVIAVILIPOS24.getCellType()==0) {CCSVIAVILIPOS24 = "" + (int) cl_CCSVIAVILIPOS24.getNumericCellValue();} else if (cl_CCSVIAVILIPOS24.getCellType()==1) {CCSVIAVILIPOS24=cl_CCSVIAVILIPOS24.getStringCellValue();}  else if (cl_CCSVIAVILIPOS24.getCellType()== 2) {CCSVIAVILIPOS24=cl_CCSVIAVILIPOS24.getRawValue();  } else { CCSVIAVILIPOS24 = "0";}}
XSSFCell cl_CCSVIAVILIPOS25_49 = worksheet.getRow(57).getCell((short) 8); if ( cl_CCSVIAVILIPOS25_49!=null){ if ( cl_CCSVIAVILIPOS25_49.getCellType()==0) {CCSVIAVILIPOS25_49 = "" + (int) cl_CCSVIAVILIPOS25_49.getNumericCellValue();} else if (cl_CCSVIAVILIPOS25_49.getCellType()==1) {CCSVIAVILIPOS25_49=cl_CCSVIAVILIPOS25_49.getStringCellValue();}  else if (cl_CCSVIAVILIPOS25_49.getCellType()== 2) {CCSVIAVILIPOS25_49=cl_CCSVIAVILIPOS25_49.getRawValue();  } else { CCSVIAVILIPOS25_49 = "0";}}
XSSFCell cl_CCSVIAVILIPOS50 = worksheet.getRow(57).getCell((short) 9); if ( cl_CCSVIAVILIPOS50!=null){ if ( cl_CCSVIAVILIPOS50.getCellType()==0) {CCSVIAVILIPOS50 = "" + (int) cl_CCSVIAVILIPOS50.getNumericCellValue();} else if (cl_CCSVIAVILIPOS50.getCellType()==1) {CCSVIAVILIPOS50=cl_CCSVIAVILIPOS50.getStringCellValue();}  else if (cl_CCSVIAVILIPOS50.getCellType()== 2) {CCSVIAVILIPOS50=cl_CCSVIAVILIPOS50.getRawValue();  } else { CCSVIAVILIPOS50 = "0";}}
XSSFCell cl_CCSCYTOLPOS24 = worksheet.getRow(58).getCell((short) 7); if ( cl_CCSCYTOLPOS24!=null){ if ( cl_CCSCYTOLPOS24.getCellType()==0) {CCSCYTOLPOS24 = "" + (int) cl_CCSCYTOLPOS24.getNumericCellValue();} else if (cl_CCSCYTOLPOS24.getCellType()==1) {CCSCYTOLPOS24=cl_CCSCYTOLPOS24.getStringCellValue();}  else if (cl_CCSCYTOLPOS24.getCellType()== 2) {CCSCYTOLPOS24=cl_CCSCYTOLPOS24.getRawValue();  } else { CCSCYTOLPOS24 = "0";}}
XSSFCell cl_CCSCYTOLPOS25_49 = worksheet.getRow(58).getCell((short) 8); if ( cl_CCSCYTOLPOS25_49!=null){ if ( cl_CCSCYTOLPOS25_49.getCellType()==0) {CCSCYTOLPOS25_49 = "" + (int) cl_CCSCYTOLPOS25_49.getNumericCellValue();} else if (cl_CCSCYTOLPOS25_49.getCellType()==1) {CCSCYTOLPOS25_49=cl_CCSCYTOLPOS25_49.getStringCellValue();}  else if (cl_CCSCYTOLPOS25_49.getCellType()== 2) {CCSCYTOLPOS25_49=cl_CCSCYTOLPOS25_49.getRawValue();  } else { CCSCYTOLPOS25_49 = "0";}}
XSSFCell cl_CCSCYTOLPOS50 = worksheet.getRow(58).getCell((short) 9); if ( cl_CCSCYTOLPOS50!=null){ if ( cl_CCSCYTOLPOS50.getCellType()==0) {CCSCYTOLPOS50 = "" + (int) cl_CCSCYTOLPOS50.getNumericCellValue();} else if (cl_CCSCYTOLPOS50.getCellType()==1) {CCSCYTOLPOS50=cl_CCSCYTOLPOS50.getStringCellValue();}  else if (cl_CCSCYTOLPOS50.getCellType()== 2) {CCSCYTOLPOS50=cl_CCSCYTOLPOS50.getRawValue();  } else { CCSCYTOLPOS50 = "0";}}
XSSFCell cl_CCSHPVPOS24 = worksheet.getRow(59).getCell((short) 7); if ( cl_CCSHPVPOS24!=null){ if ( cl_CCSHPVPOS24.getCellType()==0) {CCSHPVPOS24 = "" + (int) cl_CCSHPVPOS24.getNumericCellValue();} else if (cl_CCSHPVPOS24.getCellType()==1) {CCSHPVPOS24=cl_CCSHPVPOS24.getStringCellValue();}  else if (cl_CCSHPVPOS24.getCellType()== 2) {CCSHPVPOS24=cl_CCSHPVPOS24.getRawValue();  } else { CCSHPVPOS24 = "0";}}
XSSFCell cl_CCSHPVPOS25_49 = worksheet.getRow(59).getCell((short) 8); if ( cl_CCSHPVPOS25_49!=null){ if ( cl_CCSHPVPOS25_49.getCellType()==0) {CCSHPVPOS25_49 = "" + (int) cl_CCSHPVPOS25_49.getNumericCellValue();} else if (cl_CCSHPVPOS25_49.getCellType()==1) {CCSHPVPOS25_49=cl_CCSHPVPOS25_49.getStringCellValue();}  else if (cl_CCSHPVPOS25_49.getCellType()== 2) {CCSHPVPOS25_49=cl_CCSHPVPOS25_49.getRawValue();  } else { CCSHPVPOS25_49 = "0";}}
XSSFCell cl_CCSHPVPOS50 = worksheet.getRow(59).getCell((short) 9); if ( cl_CCSHPVPOS50!=null){ if ( cl_CCSHPVPOS50.getCellType()==0) {CCSHPVPOS50 = "" + (int) cl_CCSHPVPOS50.getNumericCellValue();} else if (cl_CCSHPVPOS50.getCellType()==1) {CCSHPVPOS50=cl_CCSHPVPOS50.getStringCellValue();}  else if (cl_CCSHPVPOS50.getCellType()== 2) {CCSHPVPOS50=cl_CCSHPVPOS50.getRawValue();  } else { CCSHPVPOS50 = "0";}}
XSSFCell cl_CCSSUSPICIOUSLES24 = worksheet.getRow(60).getCell((short) 7); if ( cl_CCSSUSPICIOUSLES24!=null){ if ( cl_CCSSUSPICIOUSLES24.getCellType()==0) {CCSSUSPICIOUSLES24 = "" + (int) cl_CCSSUSPICIOUSLES24.getNumericCellValue();} else if (cl_CCSSUSPICIOUSLES24.getCellType()==1) {CCSSUSPICIOUSLES24=cl_CCSSUSPICIOUSLES24.getStringCellValue();}  else if (cl_CCSSUSPICIOUSLES24.getCellType()== 2) {CCSSUSPICIOUSLES24=cl_CCSSUSPICIOUSLES24.getRawValue();  } else { CCSSUSPICIOUSLES24 = "0";}}
XSSFCell cl_CCSSUSPICIOUSLES25_49 = worksheet.getRow(60).getCell((short) 8); if ( cl_CCSSUSPICIOUSLES25_49!=null){ if ( cl_CCSSUSPICIOUSLES25_49.getCellType()==0) {CCSSUSPICIOUSLES25_49 = "" + (int) cl_CCSSUSPICIOUSLES25_49.getNumericCellValue();} else if (cl_CCSSUSPICIOUSLES25_49.getCellType()==1) {CCSSUSPICIOUSLES25_49=cl_CCSSUSPICIOUSLES25_49.getStringCellValue();}  else if (cl_CCSSUSPICIOUSLES25_49.getCellType()== 2) {CCSSUSPICIOUSLES25_49=cl_CCSSUSPICIOUSLES25_49.getRawValue();  } else { CCSSUSPICIOUSLES25_49 = "0";}}
XSSFCell cl_CCSSUSPICIOUSLES50 = worksheet.getRow(60).getCell((short) 9); if ( cl_CCSSUSPICIOUSLES50!=null){ if ( cl_CCSSUSPICIOUSLES50.getCellType()==0) {CCSSUSPICIOUSLES50 = "" + (int) cl_CCSSUSPICIOUSLES50.getNumericCellValue();} else if (cl_CCSSUSPICIOUSLES50.getCellType()==1) {CCSSUSPICIOUSLES50=cl_CCSSUSPICIOUSLES50.getStringCellValue();}  else if (cl_CCSSUSPICIOUSLES50.getCellType()== 2) {CCSSUSPICIOUSLES50=cl_CCSSUSPICIOUSLES50.getRawValue();  } else { CCSSUSPICIOUSLES50 = "0";}}
XSSFCell cl_CCSCryotherapy24 = worksheet.getRow(61).getCell((short) 7); if ( cl_CCSCryotherapy24!=null){ if ( cl_CCSCryotherapy24.getCellType()==0) {CCSCryotherapy24 = "" + (int) cl_CCSCryotherapy24.getNumericCellValue();} else if (cl_CCSCryotherapy24.getCellType()==1) {CCSCryotherapy24=cl_CCSCryotherapy24.getStringCellValue();}  else if (cl_CCSCryotherapy24.getCellType()== 2) {CCSCryotherapy24=cl_CCSCryotherapy24.getRawValue();  } else { CCSCryotherapy24 = "0";}}
XSSFCell cl_CCSCryotherapy25_49 = worksheet.getRow(61).getCell((short) 8); if ( cl_CCSCryotherapy25_49!=null){ if ( cl_CCSCryotherapy25_49.getCellType()==0) {CCSCryotherapy25_49 = "" + (int) cl_CCSCryotherapy25_49.getNumericCellValue();} else if (cl_CCSCryotherapy25_49.getCellType()==1) {CCSCryotherapy25_49=cl_CCSCryotherapy25_49.getStringCellValue();}  else if (cl_CCSCryotherapy25_49.getCellType()== 2) {CCSCryotherapy25_49=cl_CCSCryotherapy25_49.getRawValue();  } else { CCSCryotherapy25_49 = "0";}}
XSSFCell cl_CCSCryotherapy50 = worksheet.getRow(61).getCell((short) 9); if ( cl_CCSCryotherapy50!=null){ if ( cl_CCSCryotherapy50.getCellType()==0) {CCSCryotherapy50 = "" + (int) cl_CCSCryotherapy50.getNumericCellValue();} else if (cl_CCSCryotherapy50.getCellType()==1) {CCSCryotherapy50=cl_CCSCryotherapy50.getStringCellValue();}  else if (cl_CCSCryotherapy50.getCellType()== 2) {CCSCryotherapy50=cl_CCSCryotherapy50.getRawValue();  } else { CCSCryotherapy50 = "0";}}
XSSFCell cl_CCSLEEP24 = worksheet.getRow(62).getCell((short) 7); if ( cl_CCSLEEP24!=null){ if ( cl_CCSLEEP24.getCellType()==0) {CCSLEEP24 = "" + (int) cl_CCSLEEP24.getNumericCellValue();} else if (cl_CCSLEEP24.getCellType()==1) {CCSLEEP24=cl_CCSLEEP24.getStringCellValue();}  else if (cl_CCSLEEP24.getCellType()== 2) {CCSLEEP24=cl_CCSLEEP24.getRawValue();  } else { CCSLEEP24 = "0";}}
XSSFCell cl_CCSLEEP25_49 = worksheet.getRow(62).getCell((short) 8); if ( cl_CCSLEEP25_49!=null){ if ( cl_CCSLEEP25_49.getCellType()==0) {CCSLEEP25_49 = "" + (int) cl_CCSLEEP25_49.getNumericCellValue();} else if (cl_CCSLEEP25_49.getCellType()==1) {CCSLEEP25_49=cl_CCSLEEP25_49.getStringCellValue();}  else if (cl_CCSLEEP25_49.getCellType()== 2) {CCSLEEP25_49=cl_CCSLEEP25_49.getRawValue();  } else { CCSLEEP25_49 = "0";}}
XSSFCell cl_CCSLEEP50 = worksheet.getRow(62).getCell((short) 9); if ( cl_CCSLEEP50!=null){ if ( cl_CCSLEEP50.getCellType()==0) {CCSLEEP50 = "" + (int) cl_CCSLEEP50.getNumericCellValue();} else if (cl_CCSLEEP50.getCellType()==1) {CCSLEEP50=cl_CCSLEEP50.getStringCellValue();}  else if (cl_CCSLEEP50.getCellType()== 2) {CCSLEEP50=cl_CCSLEEP50.getRawValue();  } else { CCSLEEP50 = "0";}}
XSSFCell cl_CCSHIVPOSSCREENED24 = worksheet.getRow(63).getCell((short) 7); if ( cl_CCSHIVPOSSCREENED24!=null){ if ( cl_CCSHIVPOSSCREENED24.getCellType()==0) {CCSHIVPOSSCREENED24 = "" + (int) cl_CCSHIVPOSSCREENED24.getNumericCellValue();} else if (cl_CCSHIVPOSSCREENED24.getCellType()==1) {CCSHIVPOSSCREENED24=cl_CCSHIVPOSSCREENED24.getStringCellValue();}  else if (cl_CCSHIVPOSSCREENED24.getCellType()== 2) {CCSHIVPOSSCREENED24=cl_CCSHIVPOSSCREENED24.getRawValue();  } else { CCSHIVPOSSCREENED24 = "0";}}
XSSFCell cl_CCSHIVPOSSCREENED25_49 = worksheet.getRow(63).getCell((short) 8); if ( cl_CCSHIVPOSSCREENED25_49!=null){ if ( cl_CCSHIVPOSSCREENED25_49.getCellType()==0) {CCSHIVPOSSCREENED25_49 = "" + (int) cl_CCSHIVPOSSCREENED25_49.getNumericCellValue();} else if (cl_CCSHIVPOSSCREENED25_49.getCellType()==1) {CCSHIVPOSSCREENED25_49=cl_CCSHIVPOSSCREENED25_49.getStringCellValue();}  else if (cl_CCSHIVPOSSCREENED25_49.getCellType()== 2) {CCSHIVPOSSCREENED25_49=cl_CCSHIVPOSSCREENED25_49.getRawValue();  } else { CCSHIVPOSSCREENED25_49 = "0";}}
XSSFCell cl_CCSHIVPOSSCREENED50 = worksheet.getRow(63).getCell((short) 9); if ( cl_CCSHIVPOSSCREENED50!=null){ if ( cl_CCSHIVPOSSCREENED50.getCellType()==0) {CCSHIVPOSSCREENED50 = "" + (int) cl_CCSHIVPOSSCREENED50.getNumericCellValue();} else if (cl_CCSHIVPOSSCREENED50.getCellType()==1) {CCSHIVPOSSCREENED50=cl_CCSHIVPOSSCREENED50.getStringCellValue();}  else if (cl_CCSHIVPOSSCREENED50.getCellType()== 2) {CCSHIVPOSSCREENED50=cl_CCSHIVPOSSCREENED50.getRawValue();  } else { CCSHIVPOSSCREENED50 = "0";}}
XSSFCell cl_PNCBreastExam = worksheet.getRow(66).getCell((short) 9); if ( cl_PNCBreastExam!=null){ if ( cl_PNCBreastExam.getCellType()==0) {PNCBreastExam = "" + (int) cl_PNCBreastExam.getNumericCellValue();} else if (cl_PNCBreastExam.getCellType()==1) {PNCBreastExam=cl_PNCBreastExam.getStringCellValue();}  else if (cl_PNCBreastExam.getCellType()== 2) {PNCBreastExam=cl_PNCBreastExam.getRawValue();  } else { PNCBreastExam = "0";}}
XSSFCell cl_PNCCounselled = worksheet.getRow(67).getCell((short) 9); if ( cl_PNCCounselled!=null){ if ( cl_PNCCounselled.getCellType()==0) {PNCCounselled = "" + (int) cl_PNCCounselled.getNumericCellValue();} else if (cl_PNCCounselled.getCellType()==1) {PNCCounselled=cl_PNCCounselled.getStringCellValue();}  else if (cl_PNCCounselled.getCellType()== 2) {PNCCounselled=cl_PNCCounselled.getRawValue();  } else { PNCCounselled = "0";}}
XSSFCell cl_PNCFistula = worksheet.getRow(68).getCell((short) 9); if ( cl_PNCFistula!=null){ if ( cl_PNCFistula.getCellType()==0) {PNCFistula = "" + (int) cl_PNCFistula.getNumericCellValue();} else if (cl_PNCFistula.getCellType()==1) {PNCFistula=cl_PNCFistula.getStringCellValue();}  else if (cl_PNCFistula.getCellType()== 2) {PNCFistula=cl_PNCFistula.getRawValue();  } else { PNCFistula = "0";}}
XSSFCell cl_PNCExerNegative = worksheet.getRow(69).getCell((short) 9); if ( cl_PNCExerNegative!=null){ if ( cl_PNCExerNegative.getCellType()==0) {PNCExerNegative = "" + (int) cl_PNCExerNegative.getNumericCellValue();} else if (cl_PNCExerNegative.getCellType()==1) {PNCExerNegative=cl_PNCExerNegative.getStringCellValue();}  else if (cl_PNCExerNegative.getCellType()== 2) {PNCExerNegative=cl_PNCExerNegative.getRawValue();  } else { PNCExerNegative = "0";}}
XSSFCell cl_PNCExerPositive = worksheet.getRow(70).getCell((short) 9); if ( cl_PNCExerPositive!=null){ if ( cl_PNCExerPositive.getCellType()==0) {PNCExerPositive = "" + (int) cl_PNCExerPositive.getNumericCellValue();} else if (cl_PNCExerPositive.getCellType()==1) {PNCExerPositive=cl_PNCExerPositive.getStringCellValue();}  else if (cl_PNCExerPositive.getCellType()== 2) {PNCExerPositive=cl_PNCExerPositive.getRawValue();  } else { PNCExerPositive = "0";}}
XSSFCell cl_PNCCCSsuspect = worksheet.getRow(71).getCell((short) 9); if ( cl_PNCCCSsuspect!=null){ if ( cl_PNCCCSsuspect.getCellType()==0) {PNCCCSsuspect = "" + (int) cl_PNCCCSsuspect.getNumericCellValue();} else if (cl_PNCCCSsuspect.getCellType()==1) {PNCCCSsuspect=cl_PNCCCSsuspect.getStringCellValue();}  else if (cl_PNCCCSsuspect.getCellType()== 2) {PNCCCSsuspect=cl_PNCCCSsuspect.getRawValue();  } else { PNCCCSsuspect = "0";}}
XSSFCell cl_PNCmotherspostpartum2_3 = worksheet.getRow(72).getCell((short) 9); if ( cl_PNCmotherspostpartum2_3!=null){ if ( cl_PNCmotherspostpartum2_3.getCellType()==0) {PNCmotherspostpartum2_3 = "" + (int) cl_PNCmotherspostpartum2_3.getNumericCellValue();} else if (cl_PNCmotherspostpartum2_3.getCellType()==1) {PNCmotherspostpartum2_3=cl_PNCmotherspostpartum2_3.getStringCellValue();}  else if (cl_PNCmotherspostpartum2_3.getCellType()== 2) {PNCmotherspostpartum2_3=cl_PNCmotherspostpartum2_3.getRawValue();  } else { PNCmotherspostpartum2_3 = "0";}}
XSSFCell cl_PNCmotherspostpartum6 = worksheet.getRow(73).getCell((short) 9); if ( cl_PNCmotherspostpartum6!=null){ if ( cl_PNCmotherspostpartum6.getCellType()==0) {PNCmotherspostpartum6 = "" + (int) cl_PNCmotherspostpartum6.getNumericCellValue();} else if (cl_PNCmotherspostpartum6.getCellType()==1) {PNCmotherspostpartum6=cl_PNCmotherspostpartum6.getStringCellValue();}  else if (cl_PNCmotherspostpartum6.getCellType()== 2) {PNCmotherspostpartum6=cl_PNCmotherspostpartum6.getRawValue();  } else { PNCmotherspostpartum6 = "0";}}
XSSFCell cl_PNCinfantspostpartum2_3 = worksheet.getRow(74).getCell((short) 9); if ( cl_PNCinfantspostpartum2_3!=null){ if ( cl_PNCinfantspostpartum2_3.getCellType()==0) {PNCinfantspostpartum2_3 = "" + (int) cl_PNCinfantspostpartum2_3.getNumericCellValue();} else if (cl_PNCinfantspostpartum2_3.getCellType()==1) {PNCinfantspostpartum2_3=cl_PNCinfantspostpartum2_3.getStringCellValue();}  else if (cl_PNCinfantspostpartum2_3.getCellType()== 2) {PNCinfantspostpartum2_3=cl_PNCinfantspostpartum2_3.getRawValue();  } else { PNCinfantspostpartum2_3 = "0";}}
XSSFCell cl_PNCinfantspostpartum6 = worksheet.getRow(75).getCell((short) 9); if ( cl_PNCinfantspostpartum6!=null){ if ( cl_PNCinfantspostpartum6.getCellType()==0) {PNCinfantspostpartum6 = "" + (int) cl_PNCinfantspostpartum6.getNumericCellValue();} else if (cl_PNCinfantspostpartum6.getCellType()==1) {PNCinfantspostpartum6=cl_PNCinfantspostpartum6.getStringCellValue();}  else if (cl_PNCinfantspostpartum6.getCellType()== 2) {PNCinfantspostpartum6=cl_PNCinfantspostpartum6.getRawValue();  } else { PNCinfantspostpartum6 = "0";}}
XSSFCell cl_PNCreferralsfromotherHF = worksheet.getRow(76).getCell((short) 9); if ( cl_PNCreferralsfromotherHF!=null){ if ( cl_PNCreferralsfromotherHF.getCellType()==0) {PNCreferralsfromotherHF = "" + (int) cl_PNCreferralsfromotherHF.getNumericCellValue();} else if (cl_PNCreferralsfromotherHF.getCellType()==1) {PNCreferralsfromotherHF=cl_PNCreferralsfromotherHF.getStringCellValue();}  else if (cl_PNCreferralsfromotherHF.getCellType()== 2) {PNCreferralsfromotherHF=cl_PNCreferralsfromotherHF.getRawValue();  } else { PNCreferralsfromotherHF = "0";}}
XSSFCell cl_PNCreferralsfromotherCU = worksheet.getRow(77).getCell((short) 9); if ( cl_PNCreferralsfromotherCU!=null){ if ( cl_PNCreferralsfromotherCU.getCellType()==0) {PNCreferralsfromotherCU = "" + (int) cl_PNCreferralsfromotherCU.getNumericCellValue();} else if (cl_PNCreferralsfromotherCU.getCellType()==1) {PNCreferralsfromotherCU=cl_PNCreferralsfromotherCU.getStringCellValue();}  else if (cl_PNCreferralsfromotherCU.getCellType()== 2) {PNCreferralsfromotherCU=cl_PNCreferralsfromotherCU.getRawValue();  } else { PNCreferralsfromotherCU = "0";}}
XSSFCell cl_PNCreferralsTootherHF = worksheet.getRow(78).getCell((short) 9); if ( cl_PNCreferralsTootherHF!=null){ if ( cl_PNCreferralsTootherHF.getCellType()==0) {PNCreferralsTootherHF = "" + (int) cl_PNCreferralsTootherHF.getNumericCellValue();} else if (cl_PNCreferralsTootherHF.getCellType()==1) {PNCreferralsTootherHF=cl_PNCreferralsTootherHF.getStringCellValue();}  else if (cl_PNCreferralsTootherHF.getCellType()== 2) {PNCreferralsTootherHF=cl_PNCreferralsTootherHF.getRawValue();  } else { PNCreferralsTootherHF = "0";}}
XSSFCell cl_PNCreferralsTootherCU = worksheet.getRow(79).getCell((short) 9); if ( cl_PNCreferralsTootherCU!=null){ if ( cl_PNCreferralsTootherCU.getCellType()==0) {PNCreferralsTootherCU = "" + (int) cl_PNCreferralsTootherCU.getNumericCellValue();} else if (cl_PNCreferralsTootherCU.getCellType()==1) {PNCreferralsTootherCU=cl_PNCreferralsTootherCU.getStringCellValue();}  else if (cl_PNCreferralsTootherCU.getCellType()== 2) {PNCreferralsTootherCU=cl_PNCreferralsTootherCU.getRawValue();  } else { PNCreferralsTootherCU = "0";}}
XSSFCell cl_RsAssessed = worksheet.getRow(82).getCell((short) 9); if ( cl_RsAssessed!=null){ if ( cl_RsAssessed.getCellType()==0) {RsAssessed = "" + (int) cl_RsAssessed.getNumericCellValue();} else if (cl_RsAssessed.getCellType()==1) {RsAssessed=cl_RsAssessed.getStringCellValue();}  else if (cl_RsAssessed.getCellType()== 2) {RsAssessed=cl_RsAssessed.getRawValue();  } else { RsAssessed = "0";}}
XSSFCell cl_Rstreated = worksheet.getRow(83).getCell((short) 9); if ( cl_Rstreated!=null){ if ( cl_Rstreated.getCellType()==0) {Rstreated = "" + (int) cl_Rstreated.getNumericCellValue();} else if (cl_Rstreated.getCellType()==1) {Rstreated=cl_Rstreated.getStringCellValue();}  else if (cl_Rstreated.getCellType()== 2) {Rstreated=cl_Rstreated.getRawValue();  } else { Rstreated = "0";}}
XSSFCell cl_RsRehabilitated = worksheet.getRow(84).getCell((short) 9); if ( cl_RsRehabilitated!=null){ if ( cl_RsRehabilitated.getCellType()==0) {RsRehabilitated = "" + (int) cl_RsRehabilitated.getNumericCellValue();} else if (cl_RsRehabilitated.getCellType()==1) {RsRehabilitated=cl_RsRehabilitated.getStringCellValue();}  else if (cl_RsRehabilitated.getCellType()== 2) {RsRehabilitated=cl_RsRehabilitated.getRawValue();  } else { RsRehabilitated = "0";}}
XSSFCell cl_Rsreffered = worksheet.getRow(85).getCell((short) 9); if ( cl_Rsreffered!=null){ if ( cl_Rsreffered.getCellType()==0) {Rsreffered = "" + (int) cl_Rsreffered.getNumericCellValue();} else if (cl_Rsreffered.getCellType()==1) {Rsreffered=cl_Rsreffered.getStringCellValue();}  else if (cl_Rsreffered.getCellType()== 2) {Rsreffered=cl_Rsreffered.getRawValue();  } else { Rsreffered = "0";}}
XSSFCell cl_RsIntergrated = worksheet.getRow(86).getCell((short) 9); if ( cl_RsIntergrated!=null){ if ( cl_RsIntergrated.getCellType()==0) {RsIntergrated = "" + (int) cl_RsIntergrated.getNumericCellValue();} else if (cl_RsIntergrated.getCellType()==1) {RsIntergrated=cl_RsIntergrated.getStringCellValue();}  else if (cl_RsIntergrated.getCellType()== 2) {RsIntergrated=cl_RsIntergrated.getRawValue();  } else { RsIntergrated = "0";}}
XSSFCell cl_MSWpscounselling = worksheet.getRow(89).getCell((short) 9); if ( cl_MSWpscounselling!=null){ if ( cl_MSWpscounselling.getCellType()==0) {MSWpscounselling = "" + (int) cl_MSWpscounselling.getNumericCellValue();} else if (cl_MSWpscounselling.getCellType()==1) {MSWpscounselling=cl_MSWpscounselling.getStringCellValue();}  else if (cl_MSWpscounselling.getCellType()== 2) {MSWpscounselling=cl_MSWpscounselling.getRawValue();  } else { MSWpscounselling = "0";}}
XSSFCell cl_MSWdrugabuse = worksheet.getRow(90).getCell((short) 9); if ( cl_MSWdrugabuse!=null){ if ( cl_MSWdrugabuse.getCellType()==0) {MSWdrugabuse = "" + (int) cl_MSWdrugabuse.getNumericCellValue();} else if (cl_MSWdrugabuse.getCellType()==1) {MSWdrugabuse=cl_MSWdrugabuse.getStringCellValue();}  else if (cl_MSWdrugabuse.getCellType()== 2) {MSWdrugabuse=cl_MSWdrugabuse.getRawValue();  } else { MSWdrugabuse = "0";}}
XSSFCell cl_MSWMental = worksheet.getRow(91).getCell((short) 9); if ( cl_MSWMental!=null){ if ( cl_MSWMental.getCellType()==0) {MSWMental = "" + (int) cl_MSWMental.getNumericCellValue();} else if (cl_MSWMental.getCellType()==1) {MSWMental=cl_MSWMental.getStringCellValue();}  else if (cl_MSWMental.getCellType()== 2) {MSWMental=cl_MSWMental.getRawValue();  } else { MSWMental = "0";}}
XSSFCell cl_MSWAdolescent = worksheet.getRow(92).getCell((short) 9); if ( cl_MSWAdolescent!=null){ if ( cl_MSWAdolescent.getCellType()==0) {MSWAdolescent = "" + (int) cl_MSWAdolescent.getNumericCellValue();} else if (cl_MSWAdolescent.getCellType()==1) {MSWAdolescent=cl_MSWAdolescent.getStringCellValue();}  else if (cl_MSWAdolescent.getCellType()== 2) {MSWAdolescent=cl_MSWAdolescent.getRawValue();  } else { MSWAdolescent = "0";}}
XSSFCell cl_MSWPsAsses = worksheet.getRow(93).getCell((short) 9); if ( cl_MSWPsAsses!=null){ if ( cl_MSWPsAsses.getCellType()==0) {MSWPsAsses = "" + (int) cl_MSWPsAsses.getNumericCellValue();} else if (cl_MSWPsAsses.getCellType()==1) {MSWPsAsses=cl_MSWPsAsses.getStringCellValue();}  else if (cl_MSWPsAsses.getCellType()== 2) {MSWPsAsses=cl_MSWPsAsses.getRawValue();  } else { MSWPsAsses = "0";}}
XSSFCell cl_MSWsocialinv = worksheet.getRow(94).getCell((short) 9); if ( cl_MSWsocialinv!=null){ if ( cl_MSWsocialinv.getCellType()==0) {MSWsocialinv = "" + (int) cl_MSWsocialinv.getNumericCellValue();} else if (cl_MSWsocialinv.getCellType()==1) {MSWsocialinv=cl_MSWsocialinv.getStringCellValue();}  else if (cl_MSWsocialinv.getCellType()== 2) {MSWsocialinv=cl_MSWsocialinv.getRawValue();  } else { MSWsocialinv = "0";}}
XSSFCell cl_MSWsocialRehab = worksheet.getRow(95).getCell((short) 9); if ( cl_MSWsocialRehab!=null){ if ( cl_MSWsocialRehab.getCellType()==0) {MSWsocialRehab = "" + (int) cl_MSWsocialRehab.getNumericCellValue();} else if (cl_MSWsocialRehab.getCellType()==1) {MSWsocialRehab=cl_MSWsocialRehab.getStringCellValue();}  else if (cl_MSWsocialRehab.getCellType()== 2) {MSWsocialRehab=cl_MSWsocialRehab.getRawValue();  } else { MSWsocialRehab = "0";}}
XSSFCell cl_MSWoutreach = worksheet.getRow(96).getCell((short) 9); if ( cl_MSWoutreach!=null){ if ( cl_MSWoutreach.getCellType()==0) {MSWoutreach = "" + (int) cl_MSWoutreach.getNumericCellValue();} else if (cl_MSWoutreach.getCellType()==1) {MSWoutreach=cl_MSWoutreach.getStringCellValue();}  else if (cl_MSWoutreach.getCellType()== 2) {MSWoutreach=cl_MSWoutreach.getRawValue();  } else { MSWoutreach = "0";}}
XSSFCell cl_MSWreferrals = worksheet.getRow(97).getCell((short) 9); if ( cl_MSWreferrals!=null){ if ( cl_MSWreferrals.getCellType()==0) {MSWreferrals = "" + (int) cl_MSWreferrals.getNumericCellValue();} else if (cl_MSWreferrals.getCellType()==1) {MSWreferrals=cl_MSWreferrals.getStringCellValue();}  else if (cl_MSWreferrals.getCellType()== 2) {MSWreferrals=cl_MSWreferrals.getRawValue();  } else { MSWreferrals = "0";}}
XSSFCell cl_MSWwaivedpatients = worksheet.getRow(98).getCell((short) 9); if ( cl_MSWwaivedpatients!=null){ if ( cl_MSWwaivedpatients.getCellType()==0) {MSWwaivedpatients = "" + (int) cl_MSWwaivedpatients.getNumericCellValue();} else if (cl_MSWwaivedpatients.getCellType()==1) {MSWwaivedpatients=cl_MSWwaivedpatients.getStringCellValue();}  else if (cl_MSWwaivedpatients.getCellType()== 2) {MSWwaivedpatients=cl_MSWwaivedpatients.getRawValue();  } else { MSWwaivedpatients = "0";}}
XSSFCell cl_PsPWDOPD4 = worksheet.getRow(93).getCell((short) 14); if ( cl_PsPWDOPD4!=null){ if ( cl_PsPWDOPD4.getCellType()==0) {PsPWDOPD4 = "" + (int) cl_PsPWDOPD4.getNumericCellValue();} else if (cl_PsPWDOPD4.getCellType()==1) {PsPWDOPD4=cl_PsPWDOPD4.getStringCellValue();}  else if (cl_PsPWDOPD4.getCellType()== 2) {PsPWDOPD4=cl_PsPWDOPD4.getRawValue();  } else { PsPWDOPD4 = "0";}}
XSSFCell cl_PsPWDOPD5_19 = worksheet.getRow(93).getCell((short) 15); if ( cl_PsPWDOPD5_19!=null){ if ( cl_PsPWDOPD5_19.getCellType()==0) {PsPWDOPD5_19 = "" + (int) cl_PsPWDOPD5_19.getNumericCellValue();} else if (cl_PsPWDOPD5_19.getCellType()==1) {PsPWDOPD5_19=cl_PsPWDOPD5_19.getStringCellValue();}  else if (cl_PsPWDOPD5_19.getCellType()== 2) {PsPWDOPD5_19=cl_PsPWDOPD5_19.getRawValue();  } else { PsPWDOPD5_19 = "0";}}
XSSFCell cl_PsPWDOPD20 = worksheet.getRow(93).getCell((short) 16); if ( cl_PsPWDOPD20!=null){ if ( cl_PsPWDOPD20.getCellType()==0) {PsPWDOPD20 = "" + (int) cl_PsPWDOPD20.getNumericCellValue();} else if (cl_PsPWDOPD20.getCellType()==1) {PsPWDOPD20=cl_PsPWDOPD20.getStringCellValue();}  else if (cl_PsPWDOPD20.getCellType()== 2) {PsPWDOPD20=cl_PsPWDOPD20.getRawValue();  } else { PsPWDOPD20 = "0";}}
XSSFCell cl_PsPWDinpatient4 = worksheet.getRow(94).getCell((short) 14); if ( cl_PsPWDinpatient4!=null){ if ( cl_PsPWDinpatient4.getCellType()==0) {PsPWDinpatient4 = "" + (int) cl_PsPWDinpatient4.getNumericCellValue();} else if (cl_PsPWDinpatient4.getCellType()==1) {PsPWDinpatient4=cl_PsPWDinpatient4.getStringCellValue();}  else if (cl_PsPWDinpatient4.getCellType()== 2) {PsPWDinpatient4=cl_PsPWDinpatient4.getRawValue();  } else { PsPWDinpatient4 = "0";}}
XSSFCell cl_PsPWDinpatient5_19 = worksheet.getRow(94).getCell((short) 15); if ( cl_PsPWDinpatient5_19!=null){ if ( cl_PsPWDinpatient5_19.getCellType()==0) {PsPWDinpatient5_19 = "" + (int) cl_PsPWDinpatient5_19.getNumericCellValue();} else if (cl_PsPWDinpatient5_19.getCellType()==1) {PsPWDinpatient5_19=cl_PsPWDinpatient5_19.getStringCellValue();}  else if (cl_PsPWDinpatient5_19.getCellType()== 2) {PsPWDinpatient5_19=cl_PsPWDinpatient5_19.getRawValue();  } else { PsPWDinpatient5_19 = "0";}}
XSSFCell cl_PsPWDinpatient20 = worksheet.getRow(94).getCell((short) 16); if ( cl_PsPWDinpatient20!=null){ if ( cl_PsPWDinpatient20.getCellType()==0) {PsPWDinpatient20 = "" + (int) cl_PsPWDinpatient20.getNumericCellValue();} else if (cl_PsPWDinpatient20.getCellType()==1) {PsPWDinpatient20=cl_PsPWDinpatient20.getStringCellValue();}  else if (cl_PsPWDinpatient20.getCellType()== 2) {PsPWDinpatient20=cl_PsPWDinpatient20.getRawValue();  } else { PsPWDinpatient20 = "0";}}
XSSFCell cl_PsotherOPD4 = worksheet.getRow(95).getCell((short) 14); if ( cl_PsotherOPD4!=null){ if ( cl_PsotherOPD4.getCellType()==0) {PsotherOPD4 = "" + (int) cl_PsotherOPD4.getNumericCellValue();} else if (cl_PsotherOPD4.getCellType()==1) {PsotherOPD4=cl_PsotherOPD4.getStringCellValue();}  else if (cl_PsotherOPD4.getCellType()== 2) {PsotherOPD4=cl_PsotherOPD4.getRawValue();  } else { PsotherOPD4 = "0";}}
XSSFCell cl_PsotherOPD5_19 = worksheet.getRow(95).getCell((short) 15); if ( cl_PsotherOPD5_19!=null){ if ( cl_PsotherOPD5_19.getCellType()==0) {PsotherOPD5_19 = "" + (int) cl_PsotherOPD5_19.getNumericCellValue();} else if (cl_PsotherOPD5_19.getCellType()==1) {PsotherOPD5_19=cl_PsotherOPD5_19.getStringCellValue();}  else if (cl_PsotherOPD5_19.getCellType()== 2) {PsotherOPD5_19=cl_PsotherOPD5_19.getRawValue();  } else { PsotherOPD5_19 = "0";}}
XSSFCell cl_PsotherOPD20 = worksheet.getRow(95).getCell((short) 16); if ( cl_PsotherOPD20!=null){ if ( cl_PsotherOPD20.getCellType()==0) {PsotherOPD20 = "" + (int) cl_PsotherOPD20.getNumericCellValue();} else if (cl_PsotherOPD20.getCellType()==1) {PsotherOPD20=cl_PsotherOPD20.getStringCellValue();}  else if (cl_PsotherOPD20.getCellType()== 2) {PsotherOPD20=cl_PsotherOPD20.getRawValue();  } else { PsotherOPD20 = "0";}}
XSSFCell cl_Psotherinpatient4 = worksheet.getRow(96).getCell((short) 14); if ( cl_Psotherinpatient4!=null){ if ( cl_Psotherinpatient4.getCellType()==0) {Psotherinpatient4 = "" + (int) cl_Psotherinpatient4.getNumericCellValue();} else if (cl_Psotherinpatient4.getCellType()==1) {Psotherinpatient4=cl_Psotherinpatient4.getStringCellValue();}  else if (cl_Psotherinpatient4.getCellType()== 2) {Psotherinpatient4=cl_Psotherinpatient4.getRawValue();  } else { Psotherinpatient4 = "0";}}
XSSFCell cl_Psotherinpatient5_19 = worksheet.getRow(96).getCell((short) 15); if ( cl_Psotherinpatient5_19!=null){ if ( cl_Psotherinpatient5_19.getCellType()==0) {Psotherinpatient5_19 = "" + (int) cl_Psotherinpatient5_19.getNumericCellValue();} else if (cl_Psotherinpatient5_19.getCellType()==1) {Psotherinpatient5_19=cl_Psotherinpatient5_19.getStringCellValue();}  else if (cl_Psotherinpatient5_19.getCellType()== 2) {Psotherinpatient5_19=cl_Psotherinpatient5_19.getRawValue();  } else { Psotherinpatient5_19 = "0";}}
XSSFCell cl_Psotherinpatient20 = worksheet.getRow(96).getCell((short) 16); if ( cl_Psotherinpatient20!=null){ if ( cl_Psotherinpatient20.getCellType()==0) {Psotherinpatient20 = "" + (int) cl_Psotherinpatient20.getNumericCellValue();} else if (cl_Psotherinpatient20.getCellType()==1) {Psotherinpatient20=cl_Psotherinpatient20.getStringCellValue();}  else if (cl_Psotherinpatient20.getCellType()== 2) {Psotherinpatient20=cl_Psotherinpatient20.getRawValue();  } else { Psotherinpatient20 = "0";}}
XSSFCell cl_PsTreatments4 = worksheet.getRow(97).getCell((short) 14); if ( cl_PsTreatments4!=null){ if ( cl_PsTreatments4.getCellType()==0) {PsTreatments4 = "" + (int) cl_PsTreatments4.getNumericCellValue();} else if (cl_PsTreatments4.getCellType()==1) {PsTreatments4=cl_PsTreatments4.getStringCellValue();}  else if (cl_PsTreatments4.getCellType()== 2) {PsTreatments4=cl_PsTreatments4.getRawValue();  } else { PsTreatments4 = "0";}}
XSSFCell cl_PsTreatments5_19 = worksheet.getRow(97).getCell((short) 15); if ( cl_PsTreatments5_19!=null){ if ( cl_PsTreatments5_19.getCellType()==0) {PsTreatments5_19 = "" + (int) cl_PsTreatments5_19.getNumericCellValue();} else if (cl_PsTreatments5_19.getCellType()==1) {PsTreatments5_19=cl_PsTreatments5_19.getStringCellValue();}  else if (cl_PsTreatments5_19.getCellType()== 2) {PsTreatments5_19=cl_PsTreatments5_19.getRawValue();  } else { PsTreatments5_19 = "0";}}
XSSFCell cl_PsTreatments20 = worksheet.getRow(97).getCell((short) 16); if ( cl_PsTreatments20!=null){ if ( cl_PsTreatments20.getCellType()==0) {PsTreatments20 = "" + (int) cl_PsTreatments20.getNumericCellValue();} else if (cl_PsTreatments20.getCellType()==1) {PsTreatments20=cl_PsTreatments20.getStringCellValue();}  else if (cl_PsTreatments20.getCellType()== 2) {PsTreatments20=cl_PsTreatments20.getRawValue();  } else { PsTreatments20 = "0";}}
XSSFCell cl_PsAssessed4 = worksheet.getRow(98).getCell((short) 14); if ( cl_PsAssessed4!=null){ if ( cl_PsAssessed4.getCellType()==0) {PsAssessed4 = "" + (int) cl_PsAssessed4.getNumericCellValue();} else if (cl_PsAssessed4.getCellType()==1) {PsAssessed4=cl_PsAssessed4.getStringCellValue();}  else if (cl_PsAssessed4.getCellType()== 2) {PsAssessed4=cl_PsAssessed4.getRawValue();  } else { PsAssessed4 = "0";}}
XSSFCell cl_PsAssessed5_19 = worksheet.getRow(98).getCell((short) 15); if ( cl_PsAssessed5_19!=null){ if ( cl_PsAssessed5_19.getCellType()==0) {PsAssessed5_19 = "" + (int) cl_PsAssessed5_19.getNumericCellValue();} else if (cl_PsAssessed5_19.getCellType()==1) {PsAssessed5_19=cl_PsAssessed5_19.getStringCellValue();}  else if (cl_PsAssessed5_19.getCellType()== 2) {PsAssessed5_19=cl_PsAssessed5_19.getRawValue();  } else { PsAssessed5_19 = "0";}}
XSSFCell cl_PsAssessed20 = worksheet.getRow(98).getCell((short) 16); if ( cl_PsAssessed20!=null){ if ( cl_PsAssessed20.getCellType()==0) {PsAssessed20 = "" + (int) cl_PsAssessed20.getNumericCellValue();} else if (cl_PsAssessed20.getCellType()==1) {PsAssessed20=cl_PsAssessed20.getStringCellValue();}  else if (cl_PsAssessed20.getCellType()== 2) {PsAssessed20=cl_PsAssessed20.getRawValue();  } else { PsAssessed20 = "0";}}
XSSFCell cl_PsServices4 = worksheet.getRow(99).getCell((short) 14); if ( cl_PsServices4!=null){ if ( cl_PsServices4.getCellType()==0) {PsServices4 = "" + (int) cl_PsServices4.getNumericCellValue();} else if (cl_PsServices4.getCellType()==1) {PsServices4=cl_PsServices4.getStringCellValue();}  else if (cl_PsServices4.getCellType()== 2) {PsServices4=cl_PsServices4.getRawValue();  } else { PsServices4 = "0";}}
XSSFCell cl_PsServices5_19 = worksheet.getRow(99).getCell((short) 15); if ( cl_PsServices5_19!=null){ if ( cl_PsServices5_19.getCellType()==0) {PsServices5_19 = "" + (int) cl_PsServices5_19.getNumericCellValue();} else if (cl_PsServices5_19.getCellType()==1) {PsServices5_19=cl_PsServices5_19.getStringCellValue();}  else if (cl_PsServices5_19.getCellType()== 2) {PsServices5_19=cl_PsServices5_19.getRawValue();  } else { PsServices5_19 = "0";}}
XSSFCell cl_PsServices20 = worksheet.getRow(99).getCell((short) 16); if ( cl_PsServices20!=null){ if ( cl_PsServices20.getCellType()==0) {PsServices20 = "" + (int) cl_PsServices20.getNumericCellValue();} else if (cl_PsServices20.getCellType()==1) {PsServices20=cl_PsServices20.getStringCellValue();}  else if (cl_PsServices20.getCellType()== 2) {PsServices20=cl_PsServices20.getRawValue();  } else { PsServices20 = "0";}}
XSSFCell cl_PsANCCounsel5_19 = worksheet.getRow(100).getCell((short) 15); if ( cl_PsANCCounsel5_19!=null){ if ( cl_PsANCCounsel5_19.getCellType()==0) {PsANCCounsel5_19 = "" + (int) cl_PsANCCounsel5_19.getNumericCellValue();} else if (cl_PsANCCounsel5_19.getCellType()==1) {PsANCCounsel5_19=cl_PsANCCounsel5_19.getStringCellValue();}  else if (cl_PsANCCounsel5_19.getCellType()== 2) {PsANCCounsel5_19=cl_PsANCCounsel5_19.getRawValue();  } else { PsANCCounsel5_19 = "0";}}
XSSFCell cl_PsANCCounsel20 = worksheet.getRow(100).getCell((short) 16); if ( cl_PsANCCounsel20!=null){ if ( cl_PsANCCounsel20.getCellType()==0) {PsANCCounsel20 = "" + (int) cl_PsANCCounsel20.getNumericCellValue();} else if (cl_PsANCCounsel20.getCellType()==1) {PsANCCounsel20=cl_PsANCCounsel20.getStringCellValue();}  else if (cl_PsANCCounsel20.getCellType()== 2) {PsANCCounsel20=cl_PsANCCounsel20.getRawValue();  } else { PsANCCounsel20 = "0";}}
XSSFCell cl_PsExercise5_19 = worksheet.getRow(101).getCell((short) 15); if ( cl_PsExercise5_19!=null){ if ( cl_PsExercise5_19.getCellType()==0) {PsExercise5_19 = "" + (int) cl_PsExercise5_19.getNumericCellValue();} else if (cl_PsExercise5_19.getCellType()==1) {PsExercise5_19=cl_PsExercise5_19.getStringCellValue();}  else if (cl_PsExercise5_19.getCellType()== 2) {PsExercise5_19=cl_PsExercise5_19.getRawValue();  } else { PsExercise5_19 = "0";}}
XSSFCell cl_PsExercise20 = worksheet.getRow(101).getCell((short) 16); if ( cl_PsExercise20!=null){ if ( cl_PsExercise20.getCellType()==0) {PsExercise20 = "" + (int) cl_PsExercise20.getNumericCellValue();} else if (cl_PsExercise20.getCellType()==1) {PsExercise20=cl_PsExercise20.getStringCellValue();}  else if (cl_PsExercise20.getCellType()== 2) {PsExercise20=cl_PsExercise20.getRawValue();  } else { PsExercise20 = "0";}}
XSSFCell cl_PsFIFcollected5_19 = worksheet.getRow(102).getCell((short) 15); if ( cl_PsFIFcollected5_19!=null){ if ( cl_PsFIFcollected5_19.getCellType()==0) {PsFIFcollected5_19 = "" + (int) cl_PsFIFcollected5_19.getNumericCellValue();} else if (cl_PsFIFcollected5_19.getCellType()==1) {PsFIFcollected5_19=cl_PsFIFcollected5_19.getStringCellValue();}  else if (cl_PsFIFcollected5_19.getCellType()== 2) {PsFIFcollected5_19=cl_PsFIFcollected5_19.getRawValue();  } else { PsFIFcollected5_19 = "0";}}
XSSFCell cl_PsFIFcollected20 = worksheet.getRow(102).getCell((short) 16); if ( cl_PsFIFcollected20!=null){ if ( cl_PsFIFcollected20.getCellType()==0) {PsFIFcollected20 = "" + (int) cl_PsFIFcollected20.getNumericCellValue();} else if (cl_PsFIFcollected20.getCellType()==1) {PsFIFcollected20=cl_PsFIFcollected20.getStringCellValue();}  else if (cl_PsFIFcollected20.getCellType()== 2) {PsFIFcollected20=cl_PsFIFcollected20.getRawValue();  } else { PsFIFcollected20 = "0";}}
XSSFCell cl_PsFIFwaived5_19 = worksheet.getRow(103).getCell((short) 15); if ( cl_PsFIFwaived5_19!=null){ if ( cl_PsFIFwaived5_19.getCellType()==0) {PsFIFwaived5_19 = "" + (int) cl_PsFIFwaived5_19.getNumericCellValue();} else if (cl_PsFIFwaived5_19.getCellType()==1) {PsFIFwaived5_19=cl_PsFIFwaived5_19.getStringCellValue();}  else if (cl_PsFIFwaived5_19.getCellType()== 2) {PsFIFwaived5_19=cl_PsFIFwaived5_19.getRawValue();  } else { PsFIFwaived5_19 = "0";}}
XSSFCell cl_PsFIFwaived20 = worksheet.getRow(103).getCell((short) 16); if ( cl_PsFIFwaived20!=null){ if ( cl_PsFIFwaived20.getCellType()==0) {PsFIFwaived20 = "" + (int) cl_PsFIFwaived20.getNumericCellValue();} else if (cl_PsFIFwaived20.getCellType()==1) {PsFIFwaived20=cl_PsFIFwaived20.getStringCellValue();}  else if (cl_PsFIFwaived20.getCellType()== 2) {PsFIFwaived20=cl_PsFIFwaived20.getRawValue();  } else { PsFIFwaived20 = "0";}}
XSSFCell cl_PsFIFexempted4 = worksheet.getRow(104).getCell((short) 14); if ( cl_PsFIFexempted4!=null){ if ( cl_PsFIFexempted4.getCellType()==0) {PsFIFexempted4 = "" + (int) cl_PsFIFexempted4.getNumericCellValue();} else if (cl_PsFIFexempted4.getCellType()==1) {PsFIFexempted4=cl_PsFIFexempted4.getStringCellValue();}  else if (cl_PsFIFexempted4.getCellType()== 2) {PsFIFexempted4=cl_PsFIFexempted4.getRawValue();  } else { PsFIFexempted4 = "0";}}
XSSFCell cl_PsFIFexempted5_19 = worksheet.getRow(104).getCell((short) 15); if ( cl_PsFIFexempted5_19!=null){ if ( cl_PsFIFexempted5_19.getCellType()==0) {PsFIFexempted5_19 = "" + (int) cl_PsFIFexempted5_19.getNumericCellValue();} else if (cl_PsFIFexempted5_19.getCellType()==1) {PsFIFexempted5_19=cl_PsFIFexempted5_19.getStringCellValue();}  else if (cl_PsFIFexempted5_19.getCellType()== 2) {PsFIFexempted5_19=cl_PsFIFexempted5_19.getRawValue();  } else { PsFIFexempted5_19 = "0";}}
XSSFCell cl_PsFIFexempted20 = worksheet.getRow(104).getCell((short) 16); if ( cl_PsFIFexempted20!=null){ if ( cl_PsFIFexempted20.getCellType()==0) {PsFIFexempted20 = "" + (int) cl_PsFIFexempted20.getNumericCellValue();} else if (cl_PsFIFexempted20.getCellType()==1) {PsFIFexempted20=cl_PsFIFexempted20.getStringCellValue();}  else if (cl_PsFIFexempted20.getCellType()== 2) {PsFIFexempted20=cl_PsFIFexempted20.getRawValue();  } else { PsFIFexempted20 = "0";}}
XSSFCell cl_PsDiasbilitymeeting4 = worksheet.getRow(105).getCell((short) 14); if ( cl_PsDiasbilitymeeting4!=null){ if ( cl_PsDiasbilitymeeting4.getCellType()==0) {PsDiasbilitymeeting4 = "" + (int) cl_PsDiasbilitymeeting4.getNumericCellValue();} else if (cl_PsDiasbilitymeeting4.getCellType()==1) {PsDiasbilitymeeting4=cl_PsDiasbilitymeeting4.getStringCellValue();}  else if (cl_PsDiasbilitymeeting4.getCellType()== 2) {PsDiasbilitymeeting4=cl_PsDiasbilitymeeting4.getRawValue();  } else { PsDiasbilitymeeting4 = "0";}}
XSSFCell cl_PsDiasbilitymeeting5_19 = worksheet.getRow(105).getCell((short) 15); if ( cl_PsDiasbilitymeeting5_19!=null){ if ( cl_PsDiasbilitymeeting5_19.getCellType()==0) {PsDiasbilitymeeting5_19 = "" + (int) cl_PsDiasbilitymeeting5_19.getNumericCellValue();} else if (cl_PsDiasbilitymeeting5_19.getCellType()==1) {PsDiasbilitymeeting5_19=cl_PsDiasbilitymeeting5_19.getStringCellValue();}  else if (cl_PsDiasbilitymeeting5_19.getCellType()== 2) {PsDiasbilitymeeting5_19=cl_PsDiasbilitymeeting5_19.getRawValue();  } else { PsDiasbilitymeeting5_19 = "0";}}
XSSFCell cl_PsDiasbilitymeeting20 = worksheet.getRow(105).getCell((short) 16); if ( cl_PsDiasbilitymeeting20!=null){ if ( cl_PsDiasbilitymeeting20.getCellType()==0) {PsDiasbilitymeeting20 = "" + (int) cl_PsDiasbilitymeeting20.getNumericCellValue();} else if (cl_PsDiasbilitymeeting20.getCellType()==1) {PsDiasbilitymeeting20=cl_PsDiasbilitymeeting20.getStringCellValue();}  else if (cl_PsDiasbilitymeeting20.getCellType()== 2) {PsDiasbilitymeeting20=cl_PsDiasbilitymeeting20.getRawValue();  } else { PsDiasbilitymeeting20 = "0";}}
XSSFCell cl_bcg_u1 = worksheet.getRow(116).getCell((short) 4); if ( cl_bcg_u1!=null){ if ( cl_bcg_u1.getCellType()==0) {bcg_u1 = "" + (int) cl_bcg_u1.getNumericCellValue();} else if (cl_bcg_u1.getCellType()==1) {bcg_u1=cl_bcg_u1.getStringCellValue();}  else if (cl_bcg_u1.getCellType()== 2) {bcg_u1=cl_bcg_u1.getRawValue();  } else { bcg_u1 = "0";}}
XSSFCell cl_bcg_a1 = worksheet.getRow(117).getCell((short) 4); if ( cl_bcg_a1!=null){ if ( cl_bcg_a1.getCellType()==0) {bcg_a1 = "" + (int) cl_bcg_a1.getNumericCellValue();} else if (cl_bcg_a1.getCellType()==1) {bcg_a1=cl_bcg_a1.getStringCellValue();}  else if (cl_bcg_a1.getCellType()== 2) {bcg_a1=cl_bcg_a1.getRawValue();  } else { bcg_a1 = "0";}}
XSSFCell cl_opv_w2wk = worksheet.getRow(118).getCell((short) 4); if ( cl_opv_w2wk!=null){ if ( cl_opv_w2wk.getCellType()==0) {opv_w2wk = "" + (int) cl_opv_w2wk.getNumericCellValue();} else if (cl_opv_w2wk.getCellType()==1) {opv_w2wk=cl_opv_w2wk.getStringCellValue();}  else if (cl_opv_w2wk.getCellType()== 2) {opv_w2wk=cl_opv_w2wk.getRawValue();  } else { opv_w2wk = "0";}}
XSSFCell cl_opv1_u1 = worksheet.getRow(119).getCell((short) 4); if ( cl_opv1_u1!=null){ if ( cl_opv1_u1.getCellType()==0) {opv1_u1 = "" + (int) cl_opv1_u1.getNumericCellValue();} else if (cl_opv1_u1.getCellType()==1) {opv1_u1=cl_opv1_u1.getStringCellValue();}  else if (cl_opv1_u1.getCellType()== 2) {opv1_u1=cl_opv1_u1.getRawValue();  } else { opv1_u1 = "0";}}
XSSFCell cl_opv1_a1 = worksheet.getRow(120).getCell((short) 4); if ( cl_opv1_a1!=null){ if ( cl_opv1_a1.getCellType()==0) {opv1_a1 = "" + (int) cl_opv1_a1.getNumericCellValue();} else if (cl_opv1_a1.getCellType()==1) {opv1_a1=cl_opv1_a1.getStringCellValue();}  else if (cl_opv1_a1.getCellType()== 2) {opv1_a1=cl_opv1_a1.getRawValue();  } else { opv1_a1 = "0";}}
XSSFCell cl_opv2_u1 = worksheet.getRow(121).getCell((short) 4); if ( cl_opv2_u1!=null){ if ( cl_opv2_u1.getCellType()==0) {opv2_u1 = "" + (int) cl_opv2_u1.getNumericCellValue();} else if (cl_opv2_u1.getCellType()==1) {opv2_u1=cl_opv2_u1.getStringCellValue();}  else if (cl_opv2_u1.getCellType()== 2) {opv2_u1=cl_opv2_u1.getRawValue();  } else { opv2_u1 = "0";}}
XSSFCell cl_opv2_a1 = worksheet.getRow(122).getCell((short) 4); if ( cl_opv2_a1!=null){ if ( cl_opv2_a1.getCellType()==0) {opv2_a1 = "" + (int) cl_opv2_a1.getNumericCellValue();} else if (cl_opv2_a1.getCellType()==1) {opv2_a1=cl_opv2_a1.getStringCellValue();}  else if (cl_opv2_a1.getCellType()== 2) {opv2_a1=cl_opv2_a1.getRawValue();  } else { opv2_a1 = "0";}}
XSSFCell cl_opv3_u1 = worksheet.getRow(123).getCell((short) 4); if ( cl_opv3_u1!=null){ if ( cl_opv3_u1.getCellType()==0) {opv3_u1 = "" + (int) cl_opv3_u1.getNumericCellValue();} else if (cl_opv3_u1.getCellType()==1) {opv3_u1=cl_opv3_u1.getStringCellValue();}  else if (cl_opv3_u1.getCellType()== 2) {opv3_u1=cl_opv3_u1.getRawValue();  } else { opv3_u1 = "0";}}
XSSFCell cl_opv3_a1 = worksheet.getRow(124).getCell((short) 4); if ( cl_opv3_a1!=null){ if ( cl_opv3_a1.getCellType()==0) {opv3_a1 = "" + (int) cl_opv3_a1.getNumericCellValue();} else if (cl_opv3_a1.getCellType()==1) {opv3_a1=cl_opv3_a1.getStringCellValue();}  else if (cl_opv3_a1.getCellType()== 2) {opv3_a1=cl_opv3_a1.getRawValue();  } else { opv3_a1 = "0";}}
XSSFCell cl_ipv_u1 = worksheet.getRow(125).getCell((short) 4); if ( cl_ipv_u1!=null){ if ( cl_ipv_u1.getCellType()==0) {ipv_u1 = "" + (int) cl_ipv_u1.getNumericCellValue();} else if (cl_ipv_u1.getCellType()==1) {ipv_u1=cl_ipv_u1.getStringCellValue();}  else if (cl_ipv_u1.getCellType()== 2) {ipv_u1=cl_ipv_u1.getRawValue();  } else { ipv_u1 = "0";}}
XSSFCell cl_ipv_a1 = worksheet.getRow(126).getCell((short) 4); if ( cl_ipv_a1!=null){ if ( cl_ipv_a1.getCellType()==0) {ipv_a1 = "" + (int) cl_ipv_a1.getNumericCellValue();} else if (cl_ipv_a1.getCellType()==1) {ipv_a1=cl_ipv_a1.getStringCellValue();}  else if (cl_ipv_a1.getCellType()== 2) {ipv_a1=cl_ipv_a1.getRawValue();  } else { ipv_a1 = "0";}}
XSSFCell cl_dhh1_u1 = worksheet.getRow(127).getCell((short) 4); if ( cl_dhh1_u1!=null){ if ( cl_dhh1_u1.getCellType()==0) {dhh1_u1 = "" + (int) cl_dhh1_u1.getNumericCellValue();} else if (cl_dhh1_u1.getCellType()==1) {dhh1_u1=cl_dhh1_u1.getStringCellValue();}  else if (cl_dhh1_u1.getCellType()== 2) {dhh1_u1=cl_dhh1_u1.getRawValue();  } else { dhh1_u1 = "0";}}
XSSFCell cl_dhh1_a1 = worksheet.getRow(128).getCell((short) 4); if ( cl_dhh1_a1!=null){ if ( cl_dhh1_a1.getCellType()==0) {dhh1_a1 = "" + (int) cl_dhh1_a1.getNumericCellValue();} else if (cl_dhh1_a1.getCellType()==1) {dhh1_a1=cl_dhh1_a1.getStringCellValue();}  else if (cl_dhh1_a1.getCellType()== 2) {dhh1_a1=cl_dhh1_a1.getRawValue();  } else { dhh1_a1 = "0";}}
XSSFCell cl_dhh2_u1 = worksheet.getRow(129).getCell((short) 4); if ( cl_dhh2_u1!=null){ if ( cl_dhh2_u1.getCellType()==0) {dhh2_u1 = "" + (int) cl_dhh2_u1.getNumericCellValue();} else if (cl_dhh2_u1.getCellType()==1) {dhh2_u1=cl_dhh2_u1.getStringCellValue();}  else if (cl_dhh2_u1.getCellType()== 2) {dhh2_u1=cl_dhh2_u1.getRawValue();  } else { dhh2_u1 = "0";}}
XSSFCell cl_dhh2_a1 = worksheet.getRow(130).getCell((short) 4); if ( cl_dhh2_a1!=null){ if ( cl_dhh2_a1.getCellType()==0) {dhh2_a1 = "" + (int) cl_dhh2_a1.getNumericCellValue();} else if (cl_dhh2_a1.getCellType()==1) {dhh2_a1=cl_dhh2_a1.getStringCellValue();}  else if (cl_dhh2_a1.getCellType()== 2) {dhh2_a1=cl_dhh2_a1.getRawValue();  } else { dhh2_a1 = "0";}}
XSSFCell cl_dhh3_u1 = worksheet.getRow(131).getCell((short) 4); if ( cl_dhh3_u1!=null){ if ( cl_dhh3_u1.getCellType()==0) {dhh3_u1 = "" + (int) cl_dhh3_u1.getNumericCellValue();} else if (cl_dhh3_u1.getCellType()==1) {dhh3_u1=cl_dhh3_u1.getStringCellValue();}  else if (cl_dhh3_u1.getCellType()== 2) {dhh3_u1=cl_dhh3_u1.getRawValue();  } else { dhh3_u1 = "0";}}
XSSFCell cl_dhh3_a1 = worksheet.getRow(132).getCell((short) 4); if ( cl_dhh3_a1!=null){ if ( cl_dhh3_a1.getCellType()==0) {dhh3_a1 = "" + (int) cl_dhh3_a1.getNumericCellValue();} else if (cl_dhh3_a1.getCellType()==1) {dhh3_a1=cl_dhh3_a1.getStringCellValue();}  else if (cl_dhh3_a1.getCellType()== 2) {dhh3_a1=cl_dhh3_a1.getRawValue();  } else { dhh3_a1 = "0";}}
XSSFCell cl_pneumo1_u1 = worksheet.getRow(133).getCell((short) 4); if ( cl_pneumo1_u1!=null){ if ( cl_pneumo1_u1.getCellType()==0) {pneumo1_u1 = "" + (int) cl_pneumo1_u1.getNumericCellValue();} else if (cl_pneumo1_u1.getCellType()==1) {pneumo1_u1=cl_pneumo1_u1.getStringCellValue();}  else if (cl_pneumo1_u1.getCellType()== 2) {pneumo1_u1=cl_pneumo1_u1.getRawValue();  } else { pneumo1_u1 = "0";}}
XSSFCell cl_pneumo1_a1 = worksheet.getRow(134).getCell((short) 4); if ( cl_pneumo1_a1!=null){ if ( cl_pneumo1_a1.getCellType()==0) {pneumo1_a1 = "" + (int) cl_pneumo1_a1.getNumericCellValue();} else if (cl_pneumo1_a1.getCellType()==1) {pneumo1_a1=cl_pneumo1_a1.getStringCellValue();}  else if (cl_pneumo1_a1.getCellType()== 2) {pneumo1_a1=cl_pneumo1_a1.getRawValue();  } else { pneumo1_a1 = "0";}}
XSSFCell cl_pneumo2_u1 = worksheet.getRow(135).getCell((short) 4); if ( cl_pneumo2_u1!=null){ if ( cl_pneumo2_u1.getCellType()==0) {pneumo2_u1 = "" + (int) cl_pneumo2_u1.getNumericCellValue();} else if (cl_pneumo2_u1.getCellType()==1) {pneumo2_u1=cl_pneumo2_u1.getStringCellValue();}  else if (cl_pneumo2_u1.getCellType()== 2) {pneumo2_u1=cl_pneumo2_u1.getRawValue();  } else { pneumo2_u1 = "0";}}
XSSFCell cl_pneumo2_a1 = worksheet.getRow(136).getCell((short) 4); if ( cl_pneumo2_a1!=null){ if ( cl_pneumo2_a1.getCellType()==0) {pneumo2_a1 = "" + (int) cl_pneumo2_a1.getNumericCellValue();} else if (cl_pneumo2_a1.getCellType()==1) {pneumo2_a1=cl_pneumo2_a1.getStringCellValue();}  else if (cl_pneumo2_a1.getCellType()== 2) {pneumo2_a1=cl_pneumo2_a1.getRawValue();  } else { pneumo2_a1 = "0";}}
XSSFCell cl_pneumo3_u1 = worksheet.getRow(137).getCell((short) 4); if ( cl_pneumo3_u1!=null){ if ( cl_pneumo3_u1.getCellType()==0) {pneumo3_u1 = "" + (int) cl_pneumo3_u1.getNumericCellValue();} else if (cl_pneumo3_u1.getCellType()==1) {pneumo3_u1=cl_pneumo3_u1.getStringCellValue();}  else if (cl_pneumo3_u1.getCellType()== 2) {pneumo3_u1=cl_pneumo3_u1.getRawValue();  } else { pneumo3_u1 = "0";}}
XSSFCell cl_pneumo3_a1 = worksheet.getRow(138).getCell((short) 4); if ( cl_pneumo3_a1!=null){ if ( cl_pneumo3_a1.getCellType()==0) {pneumo3_a1 = "" + (int) cl_pneumo3_a1.getNumericCellValue();} else if (cl_pneumo3_a1.getCellType()==1) {pneumo3_a1=cl_pneumo3_a1.getStringCellValue();}  else if (cl_pneumo3_a1.getCellType()== 2) {pneumo3_a1=cl_pneumo3_a1.getRawValue();  } else { pneumo3_a1 = "0";}}
XSSFCell cl_rota1_u1 = worksheet.getRow(139).getCell((short) 4); if ( cl_rota1_u1!=null){ if ( cl_rota1_u1.getCellType()==0) {rota1_u1 = "" + (int) cl_rota1_u1.getNumericCellValue();} else if (cl_rota1_u1.getCellType()==1) {rota1_u1=cl_rota1_u1.getStringCellValue();}  else if (cl_rota1_u1.getCellType()== 2) {rota1_u1=cl_rota1_u1.getRawValue();  } else { rota1_u1 = "0";}}
XSSFCell cl_rota2_u1 = worksheet.getRow(140).getCell((short) 4); if ( cl_rota2_u1!=null){ if ( cl_rota2_u1.getCellType()==0) {rota2_u1 = "" + (int) cl_rota2_u1.getNumericCellValue();} else if (cl_rota2_u1.getCellType()==1) {rota2_u1=cl_rota2_u1.getStringCellValue();}  else if (cl_rota2_u1.getCellType()== 2) {rota2_u1=cl_rota2_u1.getRawValue();  } else { rota2_u1 = "0";}}
XSSFCell cl_vita_6 = worksheet.getRow(141).getCell((short) 4); if ( cl_vita_6!=null){ if ( cl_vita_6.getCellType()==0) {vita_6 = "" + (int) cl_vita_6.getNumericCellValue();} else if (cl_vita_6.getCellType()==1) {vita_6=cl_vita_6.getStringCellValue();}  else if (cl_vita_6.getCellType()== 2) {vita_6=cl_vita_6.getRawValue();  } else { vita_6 = "0";}}
XSSFCell cl_yv_u1 = worksheet.getRow(142).getCell((short) 4); if ( cl_yv_u1!=null){ if ( cl_yv_u1.getCellType()==0) {yv_u1 = "" + (int) cl_yv_u1.getNumericCellValue();} else if (cl_yv_u1.getCellType()==1) {yv_u1=cl_yv_u1.getStringCellValue();}  else if (cl_yv_u1.getCellType()== 2) {yv_u1=cl_yv_u1.getRawValue();  } else { yv_u1 = "0";}}
XSSFCell cl_yv_a1 = worksheet.getRow(143).getCell((short) 4); if ( cl_yv_a1!=null){ if ( cl_yv_a1.getCellType()==0) {yv_a1 = "" + (int) cl_yv_a1.getNumericCellValue();} else if (cl_yv_a1.getCellType()==1) {yv_a1=cl_yv_a1.getStringCellValue();}  else if (cl_yv_a1.getCellType()== 2) {yv_a1=cl_yv_a1.getRawValue();  } else { yv_a1 = "0";}}
XSSFCell cl_mr1_u1 = worksheet.getRow(144).getCell((short) 4); if ( cl_mr1_u1!=null){ if ( cl_mr1_u1.getCellType()==0) {mr1_u1 = "" + (int) cl_mr1_u1.getNumericCellValue();} else if (cl_mr1_u1.getCellType()==1) {mr1_u1=cl_mr1_u1.getStringCellValue();}  else if (cl_mr1_u1.getCellType()== 2) {mr1_u1=cl_mr1_u1.getRawValue();  } else { mr1_u1 = "0";}}
XSSFCell cl_mr1_a1 = worksheet.getRow(145).getCell((short) 4); if ( cl_mr1_a1!=null){ if ( cl_mr1_a1.getCellType()==0) {mr1_a1 = "" + (int) cl_mr1_a1.getNumericCellValue();} else if (cl_mr1_a1.getCellType()==1) {mr1_a1=cl_mr1_a1.getStringCellValue();}  else if (cl_mr1_a1.getCellType()== 2) {mr1_a1=cl_mr1_a1.getRawValue();  } else { mr1_a1 = "0";}}
XSSFCell cl_fic_1 = worksheet.getRow(146).getCell((short) 4); if ( cl_fic_1!=null){ if ( cl_fic_1.getCellType()==0) {fic_1 = "" + (int) cl_fic_1.getNumericCellValue();} else if (cl_fic_1.getCellType()==1) {fic_1=cl_fic_1.getStringCellValue();}  else if (cl_fic_1.getCellType()== 2) {fic_1=cl_fic_1.getRawValue();  } else { fic_1 = "0";}}
XSSFCell cl_vita_1yr = worksheet.getRow(147).getCell((short) 4); if ( cl_vita_1yr!=null){ if ( cl_vita_1yr.getCellType()==0) {vita_1yr = "" + (int) cl_vita_1yr.getNumericCellValue();} else if (cl_vita_1yr.getCellType()==1) {vita_1yr=cl_vita_1yr.getStringCellValue();}  else if (cl_vita_1yr.getCellType()== 2) {vita_1yr=cl_vita_1yr.getRawValue();  } else { vita_1yr = "0";}}
XSSFCell cl_vita_1half = worksheet.getRow(148).getCell((short) 4); if ( cl_vita_1half!=null){ if ( cl_vita_1half.getCellType()==0) {vita_1half = "" + (int) cl_vita_1half.getNumericCellValue();} else if (cl_vita_1half.getCellType()==1) {vita_1half=cl_vita_1half.getStringCellValue();}  else if (cl_vita_1half.getCellType()== 2) {vita_1half=cl_vita_1half.getRawValue();  } else { vita_1half = "0";}}
XSSFCell cl_mr2_1half = worksheet.getRow(149).getCell((short) 4); if ( cl_mr2_1half!=null){ if ( cl_mr2_1half.getCellType()==0) {mr2_1half = "" + (int) cl_mr2_1half.getNumericCellValue();} else if (cl_mr2_1half.getCellType()==1) {mr2_1half=cl_mr2_1half.getStringCellValue();}  else if (cl_mr2_1half.getCellType()== 2) {mr2_1half=cl_mr2_1half.getRawValue();  } else { mr2_1half = "0";}}
XSSFCell cl_mr2_a2 = worksheet.getRow(150).getCell((short) 4); if ( cl_mr2_a2!=null){ if ( cl_mr2_a2.getCellType()==0) {mr2_a2 = "" + (int) cl_mr2_a2.getNumericCellValue();} else if (cl_mr2_a2.getCellType()==1) {mr2_a2=cl_mr2_a2.getStringCellValue();}  else if (cl_mr2_a2.getCellType()== 2) {mr2_a2=cl_mr2_a2.getRawValue();  } else { mr2_a2 = "0";}}
XSSFCell cl_ttp_dose1 = worksheet.getRow(151).getCell((short) 4); if ( cl_ttp_dose1!=null){ if ( cl_ttp_dose1.getCellType()==0) {ttp_dose1 = "" + (int) cl_ttp_dose1.getNumericCellValue();} else if (cl_ttp_dose1.getCellType()==1) {ttp_dose1=cl_ttp_dose1.getStringCellValue();}  else if (cl_ttp_dose1.getCellType()== 2) {ttp_dose1=cl_ttp_dose1.getRawValue();  } else { ttp_dose1 = "0";}}
XSSFCell cl_ttp_dose2 = worksheet.getRow(152).getCell((short) 4); if ( cl_ttp_dose2!=null){ if ( cl_ttp_dose2.getCellType()==0) {ttp_dose2 = "" + (int) cl_ttp_dose2.getNumericCellValue();} else if (cl_ttp_dose2.getCellType()==1) {ttp_dose2=cl_ttp_dose2.getStringCellValue();}  else if (cl_ttp_dose2.getCellType()== 2) {ttp_dose2=cl_ttp_dose2.getRawValue();  } else { ttp_dose2 = "0";}}
XSSFCell cl_ttp_dose3 = worksheet.getRow(153).getCell((short) 4); if ( cl_ttp_dose3!=null){ if ( cl_ttp_dose3.getCellType()==0) {ttp_dose3 = "" + (int) cl_ttp_dose3.getNumericCellValue();} else if (cl_ttp_dose3.getCellType()==1) {ttp_dose3=cl_ttp_dose3.getStringCellValue();}  else if (cl_ttp_dose3.getCellType()== 2) {ttp_dose3=cl_ttp_dose3.getRawValue();  } else { ttp_dose3 = "0";}}
XSSFCell cl_ttp_dose4 = worksheet.getRow(154).getCell((short) 4); if ( cl_ttp_dose4!=null){ if ( cl_ttp_dose4.getCellType()==0) {ttp_dose4 = "" + (int) cl_ttp_dose4.getNumericCellValue();} else if (cl_ttp_dose4.getCellType()==1) {ttp_dose4=cl_ttp_dose4.getStringCellValue();}  else if (cl_ttp_dose4.getCellType()== 2) {ttp_dose4=cl_ttp_dose4.getRawValue();  } else { ttp_dose4 = "0";}}
XSSFCell cl_ttp_dose5 = worksheet.getRow(155).getCell((short) 4); if ( cl_ttp_dose5!=null){ if ( cl_ttp_dose5.getCellType()==0) {ttp_dose5 = "" + (int) cl_ttp_dose5.getNumericCellValue();} else if (cl_ttp_dose5.getCellType()==1) {ttp_dose5=cl_ttp_dose5.getStringCellValue();}  else if (cl_ttp_dose5.getCellType()== 2) {ttp_dose5=cl_ttp_dose5.getRawValue();  } else { ttp_dose5 = "0";}}
XSSFCell cl_ae_immun = worksheet.getRow(156).getCell((short) 4); if ( cl_ae_immun!=null){ if ( cl_ae_immun.getCellType()==0) {ae_immun = "" + (int) cl_ae_immun.getNumericCellValue();} else if (cl_ae_immun.getCellType()==1) {ae_immun=cl_ae_immun.getStringCellValue();}  else if (cl_ae_immun.getCellType()== 2) {ae_immun=cl_ae_immun.getRawValue();  } else { ae_immun = "0";}}
XSSFCell cl_vita_2_5 = worksheet.getRow(157).getCell((short) 4); if ( cl_vita_2_5!=null){ if ( cl_vita_2_5.getCellType()==0) {vita_2_5 = "" + (int) cl_vita_2_5.getNumericCellValue();} else if (cl_vita_2_5.getCellType()==1) {vita_2_5=cl_vita_2_5.getStringCellValue();}  else if (cl_vita_2_5.getCellType()== 2) {vita_2_5=cl_vita_2_5.getRawValue();  } else { vita_2_5 = "0";}}
XSSFCell cl_vita_lac_m = worksheet.getRow(158).getCell((short) 4); if ( cl_vita_lac_m!=null){ if ( cl_vita_lac_m.getCellType()==0) {vita_lac_m = "" + (int) cl_vita_lac_m.getNumericCellValue();} else if (cl_vita_lac_m.getCellType()==1) {vita_lac_m=cl_vita_lac_m.getStringCellValue();}  else if (cl_vita_lac_m.getCellType()== 2) {vita_lac_m=cl_vita_lac_m.getRawValue();  } else { vita_lac_m = "0";}}
XSSFCell cl_squint_u1 = worksheet.getRow(159).getCell((short) 4); if ( cl_squint_u1!=null){ if ( cl_squint_u1.getCellType()==0) {squint_u1 = "" + (int) cl_squint_u1.getNumericCellValue();} else if (cl_squint_u1.getCellType()==1) {squint_u1=cl_squint_u1.getStringCellValue();}  else if (cl_squint_u1.getCellType()== 2) {squint_u1=cl_squint_u1.getRawValue();  } else { squint_u1 = "0";}}
XSSFCell cl_cce_type = worksheet.getRow(163).getCell((short) 4); if ( cl_cce_type!=null){ if ( cl_cce_type.getCellType()==0) {cce_type = "" + (int) cl_cce_type.getNumericCellValue();} else if (cl_cce_type.getCellType()==1) {cce_type=cl_cce_type.getStringCellValue();}  else if (cl_cce_type.getCellType()== 2) {cce_type=cl_cce_type.getRawValue();  } else { cce_type = "0";}}
XSSFCell cl_cce_model = worksheet.getRow(164).getCell((short) 4); if ( cl_cce_model!=null){ if ( cl_cce_model.getCellType()==0) {cce_model = "" + (int) cl_cce_model.getNumericCellValue();} else if (cl_cce_model.getCellType()==1) {cce_model=cl_cce_model.getStringCellValue();}  else if (cl_cce_model.getCellType()== 2) {cce_model=cl_cce_model.getRawValue();  } else { cce_model = "0";}}
XSSFCell cl_cce_sn = worksheet.getRow(165).getCell((short) 4); if ( cl_cce_sn!=null){ if ( cl_cce_sn.getCellType()==0) {cce_sn = "" + (int) cl_cce_sn.getNumericCellValue();} else if (cl_cce_sn.getCellType()==1) {cce_sn=cl_cce_sn.getStringCellValue();}  else if (cl_cce_sn.getCellType()== 2) {cce_sn=cl_cce_sn.getRawValue();  } else { cce_sn = "0";}}
XSSFCell cl_cce_ws = worksheet.getRow(166).getCell((short) 4); if ( cl_cce_ws!=null){ if ( cl_cce_ws.getCellType()==0) {cce_ws = "" + (int) cl_cce_ws.getNumericCellValue();} else if (cl_cce_ws.getCellType()==1) {cce_ws=cl_cce_ws.getStringCellValue();}  else if (cl_cce_ws.getCellType()== 2) {cce_ws=cl_cce_ws.getRawValue();  } else { cce_ws = "0";}}
XSSFCell cl_cce_es = worksheet.getRow(167).getCell((short) 4); if ( cl_cce_es!=null){ if ( cl_cce_es.getCellType()==0) {cce_es = "" + (int) cl_cce_es.getNumericCellValue();} else if (cl_cce_es.getCellType()==1) {cce_es=cl_cce_es.getStringCellValue();}  else if (cl_cce_es.getCellType()== 2) {cce_es=cl_cce_es.getRawValue();  } else { cce_es = "0";}}
XSSFCell cl_cce_age = worksheet.getRow(168).getCell((short) 4); if ( cl_cce_age!=null){ if ( cl_cce_age.getCellType()==0) {cce_age = "" + (int) cl_cce_age.getNumericCellValue();} else if (cl_cce_age.getCellType()==1) {cce_age=cl_cce_age.getStringCellValue();}  else if (cl_cce_age.getCellType()== 2) {cce_age=cl_cce_age.getRawValue();  } else { cce_age = "0";}}
XSSFCell cl_vac_type1 = worksheet.getRow(170).getCell((short) 4); if ( cl_vac_type1!=null){ if ( cl_vac_type1.getCellType()==0) {vac_type1 = "" + (int) cl_vac_type1.getNumericCellValue();} else if (cl_vac_type1.getCellType()==1) {vac_type1=cl_vac_type1.getStringCellValue();}  else if (cl_vac_type1.getCellType()== 2) {vac_type1=cl_vac_type1.getRawValue();  } else { vac_type1 = "0";}}
XSSFCell cl_vac_days1 = worksheet.getRow(171).getCell((short) 4); if ( cl_vac_days1!=null){ if ( cl_vac_days1.getCellType()==0) {vac_days1 = "" + (int) cl_vac_days1.getNumericCellValue();} else if (cl_vac_days1.getCellType()==1) {vac_days1=cl_vac_days1.getStringCellValue();}  else if (cl_vac_days1.getCellType()== 2) {vac_days1=cl_vac_days1.getRawValue();  } else { vac_days1 = "0";}}
XSSFCell cl_vac_type2 = worksheet.getRow(172).getCell((short) 4); if ( cl_vac_type2!=null){ if ( cl_vac_type2.getCellType()==0) {vac_type2 = "" + (int) cl_vac_type2.getNumericCellValue();} else if (cl_vac_type2.getCellType()==1) {vac_type2=cl_vac_type2.getStringCellValue();}  else if (cl_vac_type2.getCellType()== 2) {vac_type2=cl_vac_type2.getRawValue();  } else { vac_type2 = "0";}}
XSSFCell cl_vac_days2 = worksheet.getRow(173).getCell((short) 4); if ( cl_vac_days2!=null){ if ( cl_vac_days2.getCellType()==0) {vac_days2 = "" + (int) cl_vac_days2.getNumericCellValue();} else if (cl_vac_days2.getCellType()==1) {vac_days2=cl_vac_days2.getStringCellValue();}  else if (cl_vac_days2.getCellType()== 2) {vac_days2=cl_vac_days2.getRawValue();  } else { vac_days2 = "0";}}
XSSFCell cl_vita_type = worksheet.getRow(174).getCell((short) 4); if ( cl_vita_type!=null){ if ( cl_vita_type.getCellType()==0) {vita_type = "" + (int) cl_vita_type.getNumericCellValue();} else if (cl_vita_type.getCellType()==1) {vita_type=cl_vita_type.getStringCellValue();}  else if (cl_vita_type.getCellType()== 2) {vita_type=cl_vita_type.getRawValue();  } else { vita_type = "0";}}
XSSFCell cl_vita_days = worksheet.getRow(175).getCell((short) 4); if ( cl_vita_days!=null){ if ( cl_vita_days.getCellType()==0) {vita_days = "" + (int) cl_vita_days.getNumericCellValue();} else if (cl_vita_days.getCellType()==1) {vita_days=cl_vita_days.getStringCellValue();}  else if (cl_vita_days.getCellType()== 2) {vita_days=cl_vita_days.getRawValue();  } else { vita_days = "0";}}
XSSFCell cl_diarrhoea = worksheet.getRow(181).getCell((short) 6); if ( cl_diarrhoea!=null){ if ( cl_diarrhoea.getCellType()==0) {diarrhoea = "" + (int) cl_diarrhoea.getNumericCellValue();} else if (cl_diarrhoea.getCellType()==1) {diarrhoea=cl_diarrhoea.getStringCellValue();}  else if (cl_diarrhoea.getCellType()== 2) {diarrhoea=cl_diarrhoea.getRawValue();  } else { diarrhoea = "0";}}
XSSFCell cl_ors_zinc = worksheet.getRow(182).getCell((short) 6); if ( cl_ors_zinc!=null){ if ( cl_ors_zinc.getCellType()==0) {ors_zinc = "" + (int) cl_ors_zinc.getNumericCellValue();} else if (cl_ors_zinc.getCellType()==1) {ors_zinc=cl_ors_zinc.getStringCellValue();}  else if (cl_ors_zinc.getCellType()== 2) {ors_zinc=cl_ors_zinc.getRawValue();  } else { ors_zinc = "0";}}
XSSFCell cl_amoxycilin = worksheet.getRow(183).getCell((short) 6); if ( cl_amoxycilin!=null){ if ( cl_amoxycilin.getCellType()==0) {amoxycilin = "" + (int) cl_amoxycilin.getNumericCellValue();} else if (cl_amoxycilin.getCellType()==1) {amoxycilin=cl_amoxycilin.getStringCellValue();}  else if (cl_amoxycilin.getCellType()== 2) {amoxycilin=cl_amoxycilin.getRawValue();  } else { amoxycilin = "0";}}
XSSFCell cl_opd_partographs = worksheet.getRow(107).getCell((short) 16); if ( cl_opd_partographs!=null){ if ( cl_opd_partographs.getCellType()==0) {opd_partographs = "" + (int) cl_opd_partographs.getNumericCellValue();} else if (cl_opd_partographs.getCellType()==1) {opd_partographs=cl_opd_partographs.getStringCellValue();}  else if (cl_opd_partographs.getCellType()== 2) {opd_partographs=cl_opd_partographs.getRawValue();  } else { opd_partographs = "0";}}
XSSFCell cl_opd_oxytocyn = worksheet.getRow(108).getCell((short) 16); if ( cl_opd_oxytocyn!=null){ if ( cl_opd_oxytocyn.getCellType()==0) {opd_oxytocyn = "" + (int) cl_opd_oxytocyn.getNumericCellValue();} else if (cl_opd_oxytocyn.getCellType()==1) {opd_oxytocyn=cl_opd_oxytocyn.getStringCellValue();}  else if (cl_opd_oxytocyn.getCellType()== 2) {opd_oxytocyn=cl_opd_oxytocyn.getRawValue();  } else { opd_oxytocyn = "0";}}
XSSFCell cl_opd_resucitated = worksheet.getRow(109).getCell((short) 16); if ( cl_opd_resucitated!=null){ if ( cl_opd_resucitated.getCellType()==0) {opd_resucitated = "" + (int) cl_opd_resucitated.getNumericCellValue();} else if (cl_opd_resucitated.getCellType()==1) {opd_resucitated=cl_opd_resucitated.getStringCellValue();}  else if (cl_opd_resucitated.getCellType()== 2) {opd_resucitated=cl_opd_resucitated.getRawValue();  } else { opd_resucitated = "0";}}

                    
                    id = facilityName + yearmonth;

                    if (!(vita_type.equals("0") && 
                            bcg_u1.equals("0") && 
                            bcg_a1.equals("0") && 
                            PMCTANCClientsT.equals("0") && 
                            (MATDeliveryT.equals("") || MATDeliveryT.equals("0") ) && 
                            FPCLIENTSN.equals("0") &&
                            FPCLIENTSR.equals("0") &&
                            ( diarrhoea.equals("") || diarrhoea.equals("0")) && 
                            CHANIS0_5TW.equals("0"))) {

                        try {
                            InsertData(conn);
                            added++;
                            
                        } catch (SQLException ex) {
                            Logger.getLogger(importdata.class.getName()).log(Level.SEVERE, null, ex);
skipped++;
                            missingFacility += "" + sheetname + " was not saved because of this error:- <b>" + ex + "</b><br>";
                        }

                    }
                    else {
                    System.out.println(" Sheet " + sheetname + " has no data");
                    missingFacility += "" + sheetname + ",";
                    skipped++;
                    }

                } else {
                    System.out.println(" Sheet " + sheetname + " has no year and month");
                    missingFacility += "" + sheetname + ",";
                    skipped++;
                }

            }//end of worksheets loop

        }//end of checking if excel file is valid
        if (conn.rs != null) {
            try {
                conn.rs.close();
            } catch (SQLException ex) {
                Logger.getLogger(importdata.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
        if (conn.st != null) {
            try {
                conn.st.close();
            } catch (SQLException ex) {
                Logger.getLogger(importdata.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
        if (conn.pst != null) {
            try {
                conn.pst.close();
            } catch (SQLException ex) {
                Logger.getLogger(importdata.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
        String unimporteddata = "";

        if (missingFacility.length() > 0) {

            unimporteddata = "<br>Data for <b>"+skipped+"<b> sites; <i>" + missingFacility.replace("'", "") + "</i><b>skipped because sheets have no year , month, Facility Name or data<b><br/>";

        }

        String updateddata = "";
        if (updated > 0) {
            updateddata = " <br/>  Data for <b>" + updated + " </b> sites <i>(" + updatedfacil + ")</i> updated <br> ";
        }

        String newdaata = "";
        if (added > 0) {

            newdaata = " Data for <b>" + added + " </b> sites imported ";
        }

        String sessionText = "<br/>" + newdaata + "   " + updateddata + " " + unimporteddata + " <br/> Importing ended at sheet no. " + lastimportedsheet;
        System.out.println("" + sessionText);
        session.setAttribute("uploadedpns", " Last uploaded file is " + fileName + " for yearmonth <b>" + yearmonth + "</b> " + sessionText);
        response.sendRedirect(nextpage);

    }

    @Override
    public String getServletInfo() {
        return "Short description";
    }// </editor-fold>

    private String getFileName(Part part) {
        String file_name = "";
        String contentDisp = part.getHeader("content-disposition");
        System.out.println("content-disposition header= " + contentDisp);
        String[] tokens = contentDisp.split(";");

        for (String token : tokens) {
            if (token.trim().startsWith("filename")) {
                file_name = token.substring(token.indexOf("=") + 2, token.length() - 1);
                break;
            }

        }
        System.out.println("content-disposition final : " + file_name);
        return file_name;
    }

    public void InsertData(dbConn conn) throws SQLException {

        String inserter = "REPLACE INTO moh711_new (id,facility,month,year,yearmonth,FPProgestinN,FPProgestinR,FPCocN,FPCocR,FPEcpN,FPEcpR,FPINJECTABLESN,FPINJECTABLESR,FPINJECTIONSN,FPINJECTIONSR,FPIUCDN,FPIUCDR,FPIMPLANTSN,FPIMPLANTSR,FPBTLN,FPBTLR,FPVasectomyN,FPVasectomyR,FPCONDOMSMN,FPCONDOMSFN,FPNaturalN,FPNaturalR,FPCLIENTSN,FPCLIENTSR,FPADOLESCENT10_14N,FPADOLESCENT10_14R,FPADOLESCENT15_19N,FPADOLESCENT15_19R,FPADOLESCENT20_24N,FPADOLESCENT20_24R,FPIUCDRemoval,FPIMPLANTSRemoval,PMCTA_1stVisit_ANC,PMCTA_ReVisit_ANC,PMCTANCClientsT,PMCTIPT1,PMCTIPT2,PMCTHB11,PMCTANCClients4,PMCTITN1,PMCTITN,PMTCTSYPHILISTES,PMTCTSYPHILISPOS,PMTCTCOUNSELLEDFEED,PMTCTBREAST,PMTCTEXERCISE,PMTCTPREG10_14,PMTCTPREG15_19,PMTCTIRON,PMTCTFOLIC,PMTCTFERROUS,MATNormalDelivery,MATCSection,MATBreech,MATAssistedVag,MATDeliveryT,MATLiveBirth,MATFreshStillBirth,MATMeceratedStillBirth,MATDeformities,MATLowAPGAR,MATWeight2500,MATTetracycline,MATPreTerm,MATDischargealive,MATbreastfeeding1,MATDeliveriesPos,MATNeoNatalD,MATMaternalD10_19,MATMaternalD,MATMaternalDAudited,MATAPHAlive,MATAPHDead,MATPPHAlive,MATPPHDead,MATEclampAlive,MATEclampDead,MATRupUtAlive,MATRupUtDead,MATObstrLaborAlive,MATObstrLaborDead,MATSepsisAlive,MATSepsisDead,MATREFFromOtherFacility,MATREFFromCU,MATREFToOtherFacility,MATREFToCU,SGBVRape72_0_9,SGBVRape72_10_17,SGBVRape72_18_49,SGBVRape72_50,SGBVinitPEP0_9,SGBVinitPEP10_17,SGBVinitPEP18_49,SGBVinitPEP50,SGBVcompPEP0_9,SGBVcompPEP10_17,SGBVcompPEP18_49,SGBVcompPEP50,SGBVPregnant0_9,SGBVPregnant10_17,SGBVPregnant18_49,SGBVPregnant50,SGBVseroconverting0_9,SGBVseroconverting10_17,SGBVseroconverting18_49,SGBVseroconverting50,SGBVsurvivors0_9,SGBVsurvivors10_17,SGBVsurvivors18_49,SGBVsurvivors50,PAC10_19,PACT,CHANIS0_5NormalweightF,CHANIS0_5NormalweightM,CHANIS0_5NormalweightT,CHANIS0_5UnderweightF,CHANIS0_5UnderweightM,CHANIS0_5UnderweightT,CHANIS0_5sevUnderweightF,CHANIS0_5sevUnderweightM,CHANIS0_5sevUnderweightT,CHANIS0_5OverweightF,CHANIS0_5OverweightM,CHANIS0_5OverweightT,CHANIS0_5ObeseF,CHANIS0_5ObeseM,CHANIS0_5ObeseT,CHANIS0_5TWF,CHANIS0_5TWM,CHANIS0_5TW,CHANIS6_23NormalweightF,CHANIS6_23NormalweightM,CHANIS6_23NormalweightT,CHANIS6_23UnderweightF,CHANIS6_23UnderweightM,CHANIS6_23UnderweightT,CHANIS6_23sevUnderweightF,CHANIS6_23sevUnderweightM,CHANIS6_23sevUnderweightT,CHANIS6_23OverweightF,CHANIS6_23OverweightM,CHANIS6_23OverweightT,CHANIS6_23ObeseF,CHANIS6_23ObeseM,CHANIS6_23ObeseT,CHANIS6_23TWF,CHANIS6_23TWM,CHANIS6_23TW,CHANIS24_59NormalweightF,CHANIS24_59NormalweightM,CHANIS24_59NormalweightT,CHANIS24_59UnderweightF,CHANIS24_59UnderweightM,CHANIS24_59UnderweightT,CHANIS24_59sevUnderweightF,CHANIS24_59sevUnderweightM,CHANIS24_59sevUnderweightT,CHANIS24_59OverweightF,CHANIS24_59OverweightM,CHANIS24_59OverweightT,CHANIS24_59ObeseF,CHANIS24_59ObeseM,CHANIS24_59ObeseT,CHANIS24_59TWF,CHANIS24_59TWM,CHANIS24_59TW,CHANISMUACNormalF,CHANISMUACNormalM,CHANISMUACNormalT,CHANISMUACModerateF,CHANISMUACModerateM,CHANISMUACModerateT,CHANISMUACSevereF,CHANISMUACSevereM,CHANISMUACSevereT,CHANISMUACMeasuredF,CHANISMUACMeasuredM,CHANISMUACMeasuredT,CHANIS0_5NormalHeightF,CHANIS0_5NormalHeightM,CHANIS0_5NormalHeightT,CHANIS0_5StuntedF,CHANIS0_5StuntedM,CHANIS0_5StuntedT,CHANIS0_5sevStuntedF,CHANIS0_5sevStuntedM,CHANIS0_5sevStuntedT,CHANIS0_5TMeasF,CHANIS0_5TMeasM,CHANIS0_5TMeas,CHANIS6_23NormalHeightF,CHANIS6_23NormalHeightM,CHANIS6_23NormalHeightT,CHANIS6_23StuntedF,CHANIS6_23StuntedM,CHANIS6_23StuntedT,CHANIS6_23sevStuntedF,CHANIS6_23sevStuntedM,CHANIS6_23sevStuntedT,CHANIS6_23TMeasF,CHANIS6_23TMeasM,CHANIS6_23TMeas,CHANIS24_59NormalHeightF,CHANIS24_59NormalHeightM,CHANIS24_59NormalHeightT,CHANIS24_59StuntedF,CHANIS24_59StuntedM,CHANIS24_59StuntedT,CHANIS24_59sevStuntedF,CHANIS24_59sevStuntedM,CHANIS24_59sevStuntedT,CHANIS24_59TMeasF,CHANIS24_59TMeasM,CHANIS24_59TMeas,CHANIS0_59NewVisitsF,CHANIS0_59NewVisitsM,CHANIS0_59NewVisitsT,CHANIS0_59KwashiakorF,CHANIS0_59KwashiakorM,CHANIS0_59KwashiakorT,CHANIS0_59MarasmusF,CHANIS0_59MarasmusM,CHANIS0_59MarasmusT,CHANIS0_59FalgrowthF,CHANIS0_59FalgrowthM,CHANIS0_59FalgrowthT,CHANIS0_59F,CHANIS0_59M,CHANIS0_59T,CHANIS0_5EXCLBreastF,CHANIS0_5EXCLBreastM,CHANIS0_5EXCLBreastT,CHANIS12_59DewormedF,CHANIS12_59DewormedM,CHANIS12_59DewormedT,CHANIS6_23MNPsF,CHANIS6_23MNPsM,CHANIS6_23MNPsT,CHANIS0_59DisabilityF,CHANIS0_59DisabilityM,CHANIS0_59DisabilityT,CCSVVH24,CCSVVH25_49,CCSVVH50,CCSPAPSMEAR24,CCSPAPSMEAR25_49,CCSPAPSMEAR50,CCSHPV24,CCSHPV25_49,CCSHPV50,CCSVIAVILIPOS24,CCSVIAVILIPOS25_49,CCSVIAVILIPOS50,CCSCYTOLPOS24,CCSCYTOLPOS25_49,CCSCYTOLPOS50,CCSHPVPOS24,CCSHPVPOS25_49,CCSHPVPOS50,CCSSUSPICIOUSLES24,CCSSUSPICIOUSLES25_49,CCSSUSPICIOUSLES50,CCSCryotherapy24,CCSCryotherapy25_49,CCSCryotherapy50,CCSLEEP24,CCSLEEP25_49,CCSLEEP50,CCSHIVPOSSCREENED24,CCSHIVPOSSCREENED25_49,CCSHIVPOSSCREENED50,PNCBreastExam,PNCCounselled,PNCFistula,PNCExerNegative,PNCExerPositive,PNCCCSsuspect,PNCmotherspostpartum2_3,PNCmotherspostpartum6,PNCinfantspostpartum2_3,PNCinfantspostpartum6,PNCreferralsfromotherHF,PNCreferralsfromotherCU,PNCreferralsTootherHF,PNCreferralsTootherCU,RsAssessed,Rstreated,RsRehabilitated,Rsreffered,RsIntergrated,MSWpscounselling,MSWdrugabuse,MSWMental,MSWAdolescent,MSWPsAsses,MSWsocialinv,MSWsocialRehab,MSWoutreach,MSWreferrals,MSWwaivedpatients,PsPWDOPD4,PsPWDOPD5_19,PsPWDOPD20,PsPWDinpatient4,PsPWDinpatient5_19,PsPWDinpatient20,PsotherOPD4,PsotherOPD5_19,PsotherOPD20,Psotherinpatient4,Psotherinpatient5_19,Psotherinpatient20,PsTreatments4,PsTreatments5_19,PsTreatments20,PsAssessed4,PsAssessed5_19,PsAssessed20,PsServices4,PsServices5_19,PsServices20,PsANCCounsel5_19,PsANCCounsel20,PsExercise5_19,PsExercise20,PsFIFcollected5_19,PsFIFcollected20,PsFIFwaived5_19,PsFIFwaived20,PsFIFexempted4,PsFIFexempted5_19,PsFIFexempted20,PsDiasbilitymeeting4,PsDiasbilitymeeting5_19,PsDiasbilitymeeting20,opd_partographs,opd_oxytocyn,opd_resucitated) "
                + "VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)";
        conn.pst = conn.connect.prepareStatement(inserter);
        conn.pst.setString(1, id);
        conn.pst.setString(2, facilityName);
        conn.pst.setString(3, reportingmonth);
        conn.pst.setString(4, reportingyear);
        conn.pst.setString(5, yearmonth);
        conn.pst.setString(6, FPProgestinN);
        conn.pst.setString(7, FPProgestinR);
        conn.pst.setString(8, FPCocN);
        conn.pst.setString(9, FPCocR);
        conn.pst.setString(10, FPEcpN);
        conn.pst.setString(11, FPEcpR);
        conn.pst.setString(12, FPINJECTABLESN);
        conn.pst.setString(13, FPINJECTABLESR);
        conn.pst.setString(14, FPINJECTIONSN);
        conn.pst.setString(15, FPINJECTIONSR);
        conn.pst.setString(16, FPIUCDN);
        conn.pst.setString(17, FPIUCDR);
        conn.pst.setString(18, FPIMPLANTSN);
        conn.pst.setString(19, FPIMPLANTSR);
        conn.pst.setString(20, FPBTLN);
        conn.pst.setString(21, FPBTLR);
        conn.pst.setString(22, FPVasectomyN);
        conn.pst.setString(23, FPVasectomyR);
        conn.pst.setString(24, FPCONDOMSMN);
        conn.pst.setString(25, FPCONDOMSFN);
        conn.pst.setString(26, FPNaturalN);
        conn.pst.setString(27, FPNaturalR);
        conn.pst.setString(28, FPCLIENTSN);
        conn.pst.setString(29, FPCLIENTSR);
        conn.pst.setString(30, FPADOLESCENT10_14N);
        conn.pst.setString(31, FPADOLESCENT10_14R);
        conn.pst.setString(32, FPADOLESCENT15_19N);
        conn.pst.setString(33, FPADOLESCENT15_19R);
        conn.pst.setString(34, FPADOLESCENT20_24N);
        conn.pst.setString(35, FPADOLESCENT20_24R);
        conn.pst.setString(36, FPIUCDRemoval);
        conn.pst.setString(37, FPIMPLANTSRemoval);
        conn.pst.setString(38, PMCTA_1stVisit_ANC);
        conn.pst.setString(39, PMCTA_ReVisit_ANC);
        conn.pst.setString(40, PMCTANCClientsT);
        conn.pst.setString(41, PMCTIPT1);
        conn.pst.setString(42, PMCTIPT2);
        conn.pst.setString(43, PMCTHB11);
        conn.pst.setString(44, PMCTANCClients4);
        conn.pst.setString(45, PMCTITN1);
        conn.pst.setString(46, PMCTITN);
        conn.pst.setString(47, PMTCTSYPHILISTES);
        conn.pst.setString(48, PMTCTSYPHILISPOS);
        conn.pst.setString(49, PMTCTCOUNSELLEDFEED);
        conn.pst.setString(50, PMTCTBREAST);
        conn.pst.setString(51, PMTCTEXERCISE);
        conn.pst.setString(52, PMTCTPREG10_14);
        conn.pst.setString(53, PMTCTPREG15_19);
        conn.pst.setString(54, PMTCTIRON);
        conn.pst.setString(55, PMTCTFOLIC);
        conn.pst.setString(56, PMTCTFERROUS);
        conn.pst.setString(57, MATNormalDelivery);
        conn.pst.setString(58, MATCSection);
        conn.pst.setString(59, MATBreech);
        conn.pst.setString(60, MATAssistedVag);
        conn.pst.setString(61, MATDeliveryT);
        conn.pst.setString(62, MATLiveBirth);
        conn.pst.setString(63, MATFreshStillBirth);
        conn.pst.setString(64, MATMeceratedStillBirth);
        conn.pst.setString(65, MATDeformities);
        conn.pst.setString(66, MATLowAPGAR);
        conn.pst.setString(67, MATWeight2500);
        conn.pst.setString(68, MATTetracycline);
        conn.pst.setString(69, MATPreTerm);
        conn.pst.setString(70, MATDischargealive);
        conn.pst.setString(71, MATbreastfeeding1);
        conn.pst.setString(72, MATDeliveriesPos);
        conn.pst.setString(73, MATNeoNatalD);
        conn.pst.setString(74, MATMaternalD10_19);
        conn.pst.setString(75, MATMaternalD);
        conn.pst.setString(76, MATMaternalDAudited);
        conn.pst.setString(77, MATAPHAlive);
        conn.pst.setString(78, MATAPHDead);
        conn.pst.setString(79, MATPPHAlive);
        conn.pst.setString(80, MATPPHDead);
        conn.pst.setString(81, MATEclampAlive);
        conn.pst.setString(82, MATEclampDead);
        conn.pst.setString(83, MATRupUtAlive);
        conn.pst.setString(84, MATRupUtDead);
        conn.pst.setString(85, MATObstrLaborAlive);
        conn.pst.setString(86, MATObstrLaborDead);
        conn.pst.setString(87, MATSepsisAlive);
        conn.pst.setString(88, MATSepsisDead);
        conn.pst.setString(89, MATREFFromOtherFacility);
        conn.pst.setString(90, MATREFFromCU);
        conn.pst.setString(91, MATREFToOtherFacility);
        conn.pst.setString(92, MATREFToCU);
        conn.pst.setString(93, SGBVRape72_0_9);
        conn.pst.setString(94, SGBVRape72_10_17);
        conn.pst.setString(95, SGBVRape72_18_49);
        conn.pst.setString(96, SGBVRape72_50);
        conn.pst.setString(97, SGBVinitPEP0_9);
        conn.pst.setString(98, SGBVinitPEP10_17);
        conn.pst.setString(99, SGBVinitPEP18_49);
        conn.pst.setString(100, SGBVinitPEP50);
        conn.pst.setString(101, SGBVcompPEP0_9);
        conn.pst.setString(102, SGBVcompPEP10_17);
        conn.pst.setString(103, SGBVcompPEP18_49);
        conn.pst.setString(104, SGBVcompPEP50);
        conn.pst.setString(105, SGBVPregnant0_9);
        conn.pst.setString(106, SGBVPregnant10_17);
        conn.pst.setString(107, SGBVPregnant18_49);
        conn.pst.setString(108, SGBVPregnant50);
        conn.pst.setString(109, SGBVseroconverting0_9);
        conn.pst.setString(110, SGBVseroconverting10_17);
        conn.pst.setString(111, SGBVseroconverting18_49);
        conn.pst.setString(112, SGBVseroconverting50);
        conn.pst.setString(113, SGBVsurvivors0_9);
        conn.pst.setString(114, SGBVsurvivors10_17);
        conn.pst.setString(115, SGBVsurvivors18_49);
        conn.pst.setString(116, SGBVsurvivors50);
        conn.pst.setString(117, PAC10_19);
        conn.pst.setString(118, PACT);
        conn.pst.setString(119, CHANIS0_5NormalweightF);
        conn.pst.setString(120, CHANIS0_5NormalweightM);
        conn.pst.setString(121, CHANIS0_5NormalweightT);
        conn.pst.setString(122, CHANIS0_5UnderweightF);
        conn.pst.setString(123, CHANIS0_5UnderweightM);
        conn.pst.setString(124, CHANIS0_5UnderweightT);
        conn.pst.setString(125, CHANIS0_5sevUnderweightF);
        conn.pst.setString(126, CHANIS0_5sevUnderweightM);
        conn.pst.setString(127, CHANIS0_5sevUnderweightT);
        conn.pst.setString(128, CHANIS0_5OverweightF);
        conn.pst.setString(129, CHANIS0_5OverweightM);
        conn.pst.setString(130, CHANIS0_5OverweightT);
        conn.pst.setString(131, CHANIS0_5ObeseF);
        conn.pst.setString(132, CHANIS0_5ObeseM);
        conn.pst.setString(133, CHANIS0_5ObeseT);
        conn.pst.setString(134, CHANIS0_5TWF);
        conn.pst.setString(135, CHANIS0_5TWM);
        conn.pst.setString(136, CHANIS0_5TW);
        conn.pst.setString(137, CHANIS6_23NormalweightF);
        conn.pst.setString(138, CHANIS6_23NormalweightM);
        conn.pst.setString(139, CHANIS6_23NormalweightT);
        conn.pst.setString(140, CHANIS6_23UnderweightF);
        conn.pst.setString(141, CHANIS6_23UnderweightM);
        conn.pst.setString(142, CHANIS6_23UnderweightT);
        conn.pst.setString(143, CHANIS6_23sevUnderweightF);
        conn.pst.setString(144, CHANIS6_23sevUnderweightM);
        conn.pst.setString(145, CHANIS6_23sevUnderweightT);
        conn.pst.setString(146, CHANIS6_23OverweightF);
        conn.pst.setString(147, CHANIS6_23OverweightM);
        conn.pst.setString(148, CHANIS6_23OverweightT);
        conn.pst.setString(149, CHANIS6_23ObeseF);
        conn.pst.setString(150, CHANIS6_23ObeseM);
        conn.pst.setString(151, CHANIS6_23ObeseT);
        conn.pst.setString(152, CHANIS6_23TWF);
        conn.pst.setString(153, CHANIS6_23TWM);
        conn.pst.setString(154, CHANIS6_23TW);
        conn.pst.setString(155, CHANIS24_59NormalweightF);
        conn.pst.setString(156, CHANIS24_59NormalweightM);
        conn.pst.setString(157, CHANIS24_59NormalweightT);
        conn.pst.setString(158, CHANIS24_59UnderweightF);
        conn.pst.setString(159, CHANIS24_59UnderweightM);
        conn.pst.setString(160, CHANIS24_59UnderweightT);
        conn.pst.setString(161, CHANIS24_59sevUnderweightF);
        conn.pst.setString(162, CHANIS24_59sevUnderweightM);
        conn.pst.setString(163, CHANIS24_59sevUnderweightT);
        conn.pst.setString(164, CHANIS24_59OverweightF);
        conn.pst.setString(165, CHANIS24_59OverweightM);
        conn.pst.setString(166, CHANIS24_59OverweightT);
        conn.pst.setString(167, CHANIS24_59ObeseF);
        conn.pst.setString(168, CHANIS24_59ObeseM);
        conn.pst.setString(169, CHANIS24_59ObeseT);
        conn.pst.setString(170, CHANIS24_59TWF);
        conn.pst.setString(171, CHANIS24_59TWM);
        conn.pst.setString(172, CHANIS24_59TW);
        conn.pst.setString(173, CHANISMUACNormalF);
        conn.pst.setString(174, CHANISMUACNormalM);
        conn.pst.setString(175, CHANISMUACNormalT);
        conn.pst.setString(176, CHANISMUACModerateF);
        conn.pst.setString(177, CHANISMUACModerateM);
        conn.pst.setString(178, CHANISMUACModerateT);
        conn.pst.setString(179, CHANISMUACSevereF);
        conn.pst.setString(180, CHANISMUACSevereM);
        conn.pst.setString(181, CHANISMUACSevereT);
        conn.pst.setString(182, CHANISMUACMeasuredF);
        conn.pst.setString(183, CHANISMUACMeasuredM);
        conn.pst.setString(184, CHANISMUACMeasuredT);
        conn.pst.setString(185, CHANIS0_5NormalHeightF);
        conn.pst.setString(186, CHANIS0_5NormalHeightM);
        conn.pst.setString(187, CHANIS0_5NormalHeightT);
        conn.pst.setString(188, CHANIS0_5StuntedF);
        conn.pst.setString(189, CHANIS0_5StuntedM);
        conn.pst.setString(190, CHANIS0_5StuntedT);
        conn.pst.setString(191, CHANIS0_5sevStuntedF);
        conn.pst.setString(192, CHANIS0_5sevStuntedM);
        conn.pst.setString(193, CHANIS0_5sevStuntedT);
        conn.pst.setString(194, CHANIS0_5TMeasF);
        conn.pst.setString(195, CHANIS0_5TMeasM);
        conn.pst.setString(196, CHANIS0_5TMeas);
        conn.pst.setString(197, CHANIS6_23NormalHeightF);
        conn.pst.setString(198, CHANIS6_23NormalHeightM);
        conn.pst.setString(199, CHANIS6_23NormalHeightT);
        conn.pst.setString(200, CHANIS6_23StuntedF);
        conn.pst.setString(201, CHANIS6_23StuntedM);
        conn.pst.setString(202, CHANIS6_23StuntedT);
        conn.pst.setString(203, CHANIS6_23sevStuntedF);
        conn.pst.setString(204, CHANIS6_23sevStuntedM);
        conn.pst.setString(205, CHANIS6_23sevStuntedT);
        conn.pst.setString(206, CHANIS6_23TMeasF);
        conn.pst.setString(207, CHANIS6_23TMeasM);
        conn.pst.setString(208, CHANIS6_23TMeas);
        conn.pst.setString(209, CHANIS24_59NormalHeightF);
        conn.pst.setString(210, CHANIS24_59NormalHeightM);
        conn.pst.setString(211, CHANIS24_59NormalHeightT);
        conn.pst.setString(212, CHANIS24_59StuntedF);
        conn.pst.setString(213, CHANIS24_59StuntedM);
        conn.pst.setString(214, CHANIS24_59StuntedT);
        conn.pst.setString(215, CHANIS24_59sevStuntedF);
        conn.pst.setString(216, CHANIS24_59sevStuntedM);
        conn.pst.setString(217, CHANIS24_59sevStuntedT);
        conn.pst.setString(218, CHANIS24_59TMeasF);
        conn.pst.setString(219, CHANIS24_59TMeasM);
        conn.pst.setString(220, CHANIS24_59TMeas);
        conn.pst.setString(221, CHANIS0_59NewVisitsF);
        conn.pst.setString(222, CHANIS0_59NewVisitsM);
        conn.pst.setString(223, CHANIS0_59NewVisitsT);
        conn.pst.setString(224, CHANIS0_59KwashiakorF);
        conn.pst.setString(225, CHANIS0_59KwashiakorM);
        conn.pst.setString(226, CHANIS0_59KwashiakorT);
        conn.pst.setString(227, CHANIS0_59MarasmusF);
        conn.pst.setString(228, CHANIS0_59MarasmusM);
        conn.pst.setString(229, CHANIS0_59MarasmusT);
        conn.pst.setString(230, CHANIS0_59FalgrowthF);
        conn.pst.setString(231, CHANIS0_59FalgrowthM);
        conn.pst.setString(232, CHANIS0_59FalgrowthT);
        conn.pst.setString(233, CHANIS0_59F);
        conn.pst.setString(234, CHANIS0_59M);
        conn.pst.setString(235, CHANIS0_59T);
        conn.pst.setString(236, CHANIS0_5EXCLBreastF);
        conn.pst.setString(237, CHANIS0_5EXCLBreastM);
        conn.pst.setString(238, CHANIS0_5EXCLBreastT);
        conn.pst.setString(239, CHANIS12_59DewormedF);
        conn.pst.setString(240, CHANIS12_59DewormedM);
        conn.pst.setString(241, CHANIS12_59DewormedT);
        conn.pst.setString(242, CHANIS6_23MNPsF);
        conn.pst.setString(243, CHANIS6_23MNPsM);
        conn.pst.setString(244, CHANIS6_23MNPsT);
        conn.pst.setString(245, CHANIS0_59DisabilityF);
        conn.pst.setString(246, CHANIS0_59DisabilityM);
        conn.pst.setString(247, CHANIS0_59DisabilityT);
        conn.pst.setString(248, CCSVVH24);
        conn.pst.setString(249, CCSVVH25_49);
        conn.pst.setString(250, CCSVVH50);
        conn.pst.setString(251, CCSPAPSMEAR24);
        conn.pst.setString(252, CCSPAPSMEAR25_49);
        conn.pst.setString(253, CCSPAPSMEAR50);
        conn.pst.setString(254, CCSHPV24);
        conn.pst.setString(255, CCSHPV25_49);
        conn.pst.setString(256, CCSHPV50);
        conn.pst.setString(257, CCSVIAVILIPOS24);
        conn.pst.setString(258, CCSVIAVILIPOS25_49);
        conn.pst.setString(259, CCSVIAVILIPOS50);
        conn.pst.setString(260, CCSCYTOLPOS24);
        conn.pst.setString(261, CCSCYTOLPOS25_49);
        conn.pst.setString(262, CCSCYTOLPOS50);
        conn.pst.setString(263, CCSHPVPOS24);
        conn.pst.setString(264, CCSHPVPOS25_49);
        conn.pst.setString(265, CCSHPVPOS50);
        conn.pst.setString(266, CCSSUSPICIOUSLES24);
        conn.pst.setString(267, CCSSUSPICIOUSLES25_49);
        conn.pst.setString(268, CCSSUSPICIOUSLES50);
        conn.pst.setString(269, CCSCryotherapy24);
        conn.pst.setString(270, CCSCryotherapy25_49);
        conn.pst.setString(271, CCSCryotherapy50);
        conn.pst.setString(272, CCSLEEP24);
        conn.pst.setString(273, CCSLEEP25_49);
        conn.pst.setString(274, CCSLEEP50);
        conn.pst.setString(275, CCSHIVPOSSCREENED24);
        conn.pst.setString(276, CCSHIVPOSSCREENED25_49);
        conn.pst.setString(277, CCSHIVPOSSCREENED50);
        conn.pst.setString(278, PNCBreastExam);
        conn.pst.setString(279, PNCCounselled);
        conn.pst.setString(280, PNCFistula);
        conn.pst.setString(281, PNCExerNegative);
        conn.pst.setString(282, PNCExerPositive);
        conn.pst.setString(283, PNCCCSsuspect);
        conn.pst.setString(284, PNCmotherspostpartum2_3);
        conn.pst.setString(285, PNCmotherspostpartum6);
        conn.pst.setString(286, PNCinfantspostpartum2_3);
        conn.pst.setString(287, PNCinfantspostpartum6);
        conn.pst.setString(288, PNCreferralsfromotherHF);
        conn.pst.setString(289, PNCreferralsfromotherCU);
        conn.pst.setString(290, PNCreferralsTootherHF);
        conn.pst.setString(291, PNCreferralsTootherCU);
        conn.pst.setString(292, RsAssessed);
        conn.pst.setString(293, Rstreated);
        conn.pst.setString(294, RsRehabilitated);
        conn.pst.setString(295, Rsreffered);
        conn.pst.setString(296, RsIntergrated);
        conn.pst.setString(297, MSWpscounselling);
        conn.pst.setString(298, MSWdrugabuse);
        conn.pst.setString(299, MSWMental);
        conn.pst.setString(300, MSWAdolescent);
        conn.pst.setString(301, MSWPsAsses);
        conn.pst.setString(302, MSWsocialinv);
        conn.pst.setString(303, MSWsocialRehab);
        conn.pst.setString(304, MSWoutreach);
        conn.pst.setString(305, MSWreferrals);
        conn.pst.setString(306, MSWwaivedpatients);
        conn.pst.setString(307, PsPWDOPD4);
        conn.pst.setString(308, PsPWDOPD5_19);
        conn.pst.setString(309, PsPWDOPD20);
        conn.pst.setString(310, PsPWDinpatient4);
        conn.pst.setString(311, PsPWDinpatient5_19);
        conn.pst.setString(312, PsPWDinpatient20);
        conn.pst.setString(313, PsotherOPD4);
        conn.pst.setString(314, PsotherOPD5_19);
        conn.pst.setString(315, PsotherOPD20);
        conn.pst.setString(316, Psotherinpatient4);
        conn.pst.setString(317, Psotherinpatient5_19);
        conn.pst.setString(318, Psotherinpatient20);
        conn.pst.setString(319, PsTreatments4);
        conn.pst.setString(320, PsTreatments5_19);
        conn.pst.setString(321, PsTreatments20);
        conn.pst.setString(322, PsAssessed4);
        conn.pst.setString(323, PsAssessed5_19);
        conn.pst.setString(324, PsAssessed20);
        conn.pst.setString(325, PsServices4);
        conn.pst.setString(326, PsServices5_19);
        conn.pst.setString(327, PsServices20);
        conn.pst.setString(328, PsANCCounsel5_19);
        conn.pst.setString(329, PsANCCounsel20);
        conn.pst.setString(330, PsExercise5_19);
        conn.pst.setString(331, PsExercise20);
        conn.pst.setString(332, PsFIFcollected5_19);
        conn.pst.setString(333, PsFIFcollected20);
        conn.pst.setString(334, PsFIFwaived5_19);
        conn.pst.setString(335, PsFIFwaived20);
        conn.pst.setString(336, PsFIFexempted4);
        conn.pst.setString(337, PsFIFexempted5_19);
        conn.pst.setString(338, PsFIFexempted20);
        conn.pst.setString(339, PsDiasbilitymeeting4);
        conn.pst.setString(340, PsDiasbilitymeeting5_19);
        conn.pst.setString(341, PsDiasbilitymeeting20);
        conn.pst.setString(342, opd_partographs);
        conn.pst.setString(343, opd_oxytocyn);
        conn.pst.setString(344, opd_resucitated);
        System.out.println("moh731_" + facilityName + "_" + conn.pst.executeUpdate());
                       // System.out.println(""+conn.pst);

        if (cce_age.length() > 5) {
            cce_age = cce_age.substring(0, 5);
        }

        String inserter1 = "REPLACE INTO moh710 (id,facility,month,year,yearmonth,bcg_u1,bcg_a1,opv_w2wk,opv1_u1,opv1_a1,opv2_u1,opv2_a1,opv3_u1,opv3_a1,ipv_u1,ipv_a1,dhh1_u1,dhh1_a1,dhh2_u1,dhh2_a1,dhh3_u1,dhh3_a1,pneumo1_u1,pneumo1_a1,pneumo2_u1,pneumo2_a1,pneumo3_u1,pneumo3_a1,rota1_u1,rota2_u1,vita_6,yv_u1,yv_a1,mr1_u1,mr1_a1,fic_1,vita_1yr,vita_1half,mr2_1half,mr2_a2,ttp_dose1,ttp_dose2,ttp_dose3,ttp_dose4,ttp_dose5,ae_immun,vita_2_5,vita_lac_m,squint_u1,cce_type,cce_model,cce_sn,cce_ws,cce_es,cce_age,vac_type1,vac_days1,vac_type2,vac_days2,vita_type,vita_days,diarrhoea,ors_zinc,amoxycilin) Values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)";
        conn.pst = conn.connect.prepareStatement(inserter1);
        conn.pst.setString(1, id);
        conn.pst.setString(2, facilityName);
        conn.pst.setString(3, reportingmonth);
        conn.pst.setString(4, reportingyear);
        conn.pst.setString(5, yearmonth);
        conn.pst.setString(6, bcg_u1);
        conn.pst.setString(7, bcg_a1);
        conn.pst.setString(8, opv_w2wk);
        conn.pst.setString(9, opv1_u1);
        conn.pst.setString(10, opv1_a1);
        conn.pst.setString(11, opv2_u1);
        conn.pst.setString(12, opv2_a1);
        conn.pst.setString(13, opv3_u1);
        conn.pst.setString(14, opv3_a1);
        conn.pst.setString(15, ipv_u1);
        conn.pst.setString(16, ipv_a1);
        conn.pst.setString(17, dhh1_u1);
        conn.pst.setString(18, dhh1_a1);
        conn.pst.setString(19, dhh2_u1);
        conn.pst.setString(20, dhh2_a1);
        conn.pst.setString(21, dhh3_u1);
        conn.pst.setString(22, dhh3_a1);
        conn.pst.setString(23, pneumo1_u1);
        conn.pst.setString(24, pneumo1_a1);
        conn.pst.setString(25, pneumo2_u1);
        conn.pst.setString(26, pneumo2_a1);
        conn.pst.setString(27, pneumo3_u1);
        conn.pst.setString(28, pneumo3_a1);
        conn.pst.setString(29, rota1_u1);
        conn.pst.setString(30, rota2_u1);
        conn.pst.setString(31, vita_6);
        conn.pst.setString(32, yv_u1);
        conn.pst.setString(33, yv_a1);
        conn.pst.setString(34, mr1_u1);
        conn.pst.setString(35, mr1_a1);
        conn.pst.setString(36, fic_1);
        conn.pst.setString(37, vita_1yr);
        conn.pst.setString(38, vita_1half);
        conn.pst.setString(39, mr2_1half);
        conn.pst.setString(40, mr2_a2);
        conn.pst.setString(41, ttp_dose1);
        conn.pst.setString(42, ttp_dose2);
        conn.pst.setString(43, ttp_dose3);
        conn.pst.setString(44, ttp_dose4);
        conn.pst.setString(45, ttp_dose5);
        conn.pst.setString(46, ae_immun);
        conn.pst.setString(47, vita_2_5);
        conn.pst.setString(48, vita_lac_m);
        conn.pst.setString(49, squint_u1);
        conn.pst.setString(50, cce_type);
        conn.pst.setString(51, cce_model);
        conn.pst.setString(52, cce_sn);
        conn.pst.setString(53, cce_ws);
        conn.pst.setString(54, cce_es);
        conn.pst.setString(55, cce_age);
        conn.pst.setString(56, vac_type1);
        conn.pst.setString(57, vac_days1);
        conn.pst.setString(58, vac_type2);
        conn.pst.setString(59, vac_days2);
        conn.pst.setString(60, vita_type);
        conn.pst.setString(61, vita_days);
        conn.pst.setString(62, diarrhoea);
        conn.pst.setString(63, ors_zinc);
        conn.pst.setString(64, amoxycilin);

        conn.pst.executeUpdate();

                         //System.out.println(cce_age.length()+" cceage "+cce_age+"_"+facilityName);
        //System.out.println(""+conn.pst);
        // System.out.println("moh710_"+facilityName+"_"+conn.pst.executeUpdate());
    }

}
