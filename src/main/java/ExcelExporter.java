
import jxl.Workbook;
import jxl.write.*;
import jxl.write.Number;

import java.io.*;
import java.net.URL;
import java.util.*;


public class ExcelExporter {

    private static class sprintObject implements Comparable<sprintObject>{
        String issueType,issueKey,issueSummary, storyPts;

        public sprintObject(String issue, String key, String summary){
            this.issueType = issue;
            this.issueKey = key;
            this.issueSummary = summary;
//            this.storyPts = story;
        }

        public String getStoryPts(){return this.storyPts;}

        public void setStoryPts(String story){this.storyPts = story;}

        public String getIssueType() {
            return issueType;
        }

        public void setIssueType(String issueType) {
            this.issueType = issueType;
        }

        public String getIssueKey() {
            return issueKey;
        }

        public void setIssueKey(String issueKey) {
            this.issueKey = issueKey;
        }

        public String getIssueSummary() {
            return issueSummary;
        }

        public void setIssueSummary(String issueSummary) {
            this.issueSummary = issueSummary;
        }

        @Override
        public String toString(){
            return this.issueType + " " + this.issueKey + " " + this.issueSummary;
        }

        public int compareTo(sprintObject compareFruit) {

            return this.issueKey.compareTo(compareFruit.issueKey);
        }
    }


    private static String _excelFileLocation = System.getProperty("user.dir");//+"/temp.xls";
    private static List<sprintObject> listOfItems = new ArrayList<sprintObject>();
    private static WritableWorkbook myFirstWbook = null;
    private static String baseJiraUrl = "https://medbridge.atlassian.net/browse/";
    static String sprintName = null;
    static String locOfCSV = null;

    private static void setupArgs(String [] args) throws Exception{

        if(args.length < 1)
            throw new Exception("You must have the following arguments\n" +
                    "arg[0] = Location of CSV From JIRA\n" +
                    "arg[1] = Location of where you want to have the file stored\n" +
                    "\n Ex. java -jar jiraToSheets.jar ~/user/anonymous/jiraFile.csv ~/user/anonymous/Desktop/SomeExcelFile.xls"
            );


        if (args.length == 2) {
            locOfCSV = args[0];
            sprintName = args[1];
        }
    }

    public static void main(String [] args) throws Exception{


        setupArgs(args);

//
//        locOfCSV = "Alpha_Sprint_L.csv";
//        sprintName = "Sprint_L_Name";
//        locOfCSV = locOfCSV;
//        sprintName = "Sprint_J_Name";
        _excelFileLocation = _excelFileLocation + "/" + sprintName + ".x";
        System.out.println(_excelFileLocation);

        readCSVFile(locOfCSV);

        //Get that ABC order going
        Collections.sort(listOfItems);

        try {
            myFirstWbook = Workbook.createWorkbook(new File(_excelFileLocation));

            //create the overview test plan
            WritableSheet excelSheet = myFirstWbook.createSheet("Test_Plans", 0);

            addHeaders();
//
            int counter = 1;
            for(sprintObject s : listOfItems){
                createIndividualSheets(s,counter);
                //myFirstWbook.createSheet(s.getIssueKey(),counter);
                counter++;
            }

            String currTicketType = "";
            for(int x = 1; x < listOfItems.size()+1;x++) {

                //Get Hyperlink and labels
                Label labelType = new Label(0, x, listOfItems.get(x-1).getIssueType());
                WritableHyperlink link = (new WritableHyperlink(1,x,new URL(baseJiraUrl+listOfItems.get(x-1).getIssueKey())));
                link.setDescription(listOfItems.get(x-1).getIssueKey());
                Label labelSummary = new Label(2, x, listOfItems.get(x-1).getIssueSummary());
                Label labelStoryPt = new Label(4,x,listOfItems.get(x-1).getStoryPts());

                //Add to sheet
                excelSheet.addCell(labelType);
                excelSheet.addHyperlink(link);
                excelSheet.addCell(labelSummary);
                excelSheet.addCell(labelStoryPt);
            }

            myFirstWbook.write();


        } catch (Exception e) {
            System.out.println("didnt' make it " + e.getMessage());
        } finally {

            if (myFirstWbook != null) {
                try {
                    myFirstWbook.close();
                } catch (IOException e) {
                    e.printStackTrace();
                } catch (WriteException e) {
                    e.printStackTrace();
                }//catch
            }//if not null


        }//finally
    }//main

    private static void touchFile(){

        String command = System.getProperty("user.dir");
            //Needed for pinging waiting for device to come back on
            if(!command.contains("ping"))
                command = System.getProperty("user.home")+"/Library/Android/sdk/platform-tools/adb " + command;


            StringBuffer output = new StringBuffer();

            Process p;
            try {
                p = Runtime.getRuntime().exec(command);
                p.waitFor();
                BufferedReader reader =
                        new BufferedReader(new InputStreamReader(p.getInputStream()));

                String line = "";
                while ((line = reader.readLine())!= null) {
                    if(line.contains("10.")) {
                        String [] arr = line.split("src");
                        output.append(arr[1]);
                        break;
                    }else{
                        output.append(line + "\n");
                    }

                }

            } catch (Exception e) {
                e.printStackTrace();
            }

    }

    //TODO: Need to finish getting all of the naming conventions
    private String getJiraType(String type){

        if(type.equals("ACA"))
            return "Android";

        if(type.equals("REP"))
            return "Reporting/Analytics";

        if(type.equals("ACC"))
            return "Accreditation Updates";

        //if(type.equals("Con"))

        return null;
    }

    private static void readCSVFile(String fileLoc){
        Scanner scan;

        try{
           scan  = new Scanner(new File(fileLoc));
           //Dump the header line
           scan.nextLine();

            //on the full line 0 is summary, 1 = issue key, 4 = issue type, 59 = story pts
           while(scan.hasNextLine()){
               String [] arrayParse = scan.nextLine().split(",");

               //System.out.println(arrayParse.toString());
               if(arrayParse[0].equalsIgnoreCase("story")||arrayParse[0].equalsIgnoreCase("bug")){
                   try {
                       listOfItems.add(new sprintObject(arrayParse[0], arrayParse[1], arrayParse[3]));//, arrayParse[11]));
                       System.out.println(new sprintObject(arrayParse[0], arrayParse[1], arrayParse[3]).toString());//, arrayParse[11]).toString());
                   }catch(Exception e){
                        System.out.println(e.getMessage());
                   }
               }
           }
        }catch(FileNotFoundException e){
            System.out.print("File not found exception...\n" + e.getStackTrace());
        }
    }

    private static void createIndividualSheets(sprintObject obj, int count){
/*
        //Hyperinks https://stackoverflow.com/questions/16195140/how-do-i-activate-a-hyperlink-in-excel-after-writing-it-in-jexcel
 */
        String [] headers = {
                "Test Case Name:","Description:", "Test Case Completed Date:",
                "Run By:","Start Date","Finish Date","Jira Ticket","Time(How long did it take","Environment",
                "Build #","Prerequisite","Os / Browser:", "Assumptions","Overall Pass or Fail"
        };

        String [] workHeaders = {
               "Title","Action","Expected Result","Actual Results", "Pass / Fail", "Notes"
        };

        Label label;
        WritableSheet excelSheet = myFirstWbook.createSheet(obj.getIssueKey(),count);

        String jiraLinkURL = baseJiraUrl+obj.getIssueKey();
        String linkDesc = obj.getIssueKey();
        WritableCellFormat header = new WritableCellFormat();

        try {

            header.setBackground(Colour.LIGHT_GREEN);

            for (int x = 0; x < headers.length; x++) {
                label = new Label(0, x, headers[x]);
                excelSheet.addCell(label);

                //Test Case Name
                if(headers[x].equalsIgnoreCase("Test Case Name:")){
                    label = (new Label(1,x,obj.getIssueKey()));
                    excelSheet.addCell(label);
                }

                //Description
                if(headers[x].equalsIgnoreCase("Description:")){
                    label = (new Label(1,x,obj.getIssueSummary()));
                    excelSheet.addCell(label);
                }

                //Hyper Link
                if (headers[x].equalsIgnoreCase("Jira Ticket")) {
                    WritableHyperlink link = (new WritableHyperlink(1,x,new URL(jiraLinkURL)));
                    link.setDescription(linkDesc);
                    excelSheet.addHyperlink(link);
                }


            }//end of for loop

            for(int x = 0 ; x < workHeaders.length;x++) {
                label = new Label(x, 15, workHeaders[x]);
                excelSheet.addCell(label);
                WritableCell c = excelSheet.getWritableCell(x,15);
                c.setCellFormat(header);
            }


        }catch (Exception e){
            System.out.print("Error in creating individual worksheets " + e.getStackTrace());
        }


    }

    private static void addHeaders(){
        String [] arrayHeaders ={"Issue Type","Jira Ticket", "Summary", "QA Owner","Story Points", "Result(P/F)","Notes"};
        WritableCellFormat header = new WritableCellFormat();

        WritableSheet excelSheet = myFirstWbook.getSheet(0);

        try {

            header.setBackground(Colour.LIGHT_TURQUOISE);

            for(int x = 0; x < arrayHeaders.length; x ++) {

                Label label = new Label(x,0,arrayHeaders[x]);
                excelSheet.addCell(label);
                WritableCell c = excelSheet.getWritableCell(x,0);
                c.setCellFormat(header);
            }

            //myFirstWbook.write();

        }catch(Exception e){
            System.out.print("Error thrown while trying to do the headers. (addHeaders Method)" + e.getStackTrace());
        }

    }

}
