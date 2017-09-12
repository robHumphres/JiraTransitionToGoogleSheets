
import jxl.Workbook;
import jxl.write.*;
import jxl.write.Number;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.net.URL;
import java.util.*;


public class ExcelExporter {

    private static class sprintObject implements Comparable<sprintObject>{
        String issueType,issueKey,issueSummary;


        public sprintObject(String issue, String key, String summary){
            this.issueType = issue;
            this.issueKey = key;
            this.issueSummary = summary;
        }

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


    private static String _excelFileLocation = "/Users/robert/Desktop/temp.xls";//System.getProperty("user.dir")+"temp.xls";
    private static List<sprintObject> listOfItems = new ArrayList<sprintObject>();
    private static WritableWorkbook myFirstWbook = null;

    private static void readCSVFile(String fileLoc){
        Scanner scan;

        try{
           scan  = new Scanner(new File(fileLoc));
           //Dump the header line
           scan.nextLine();


           while(scan.hasNextLine()){
               String [] arrayParse = scan.nextLine().split(",");
               if(arrayParse[0].equalsIgnoreCase("story")){
                   listOfItems.add(new sprintObject(arrayParse[0],arrayParse[1],arrayParse[4]));
               }
           }
        }catch(FileNotFoundException e){
            System.out.print("File not found exception...\n" + e.getStackTrace());
        }
    }


    public static void main(String [] args) {
        String sprintName;
        String locOfCSV = null;

        if(args.length == 1)
            locOfCSV = args[0];

        if (args.length == 2) {
            sprintName = args[1];
            _excelFileLocation = System.getProperty("user.dir") + sprintName + ".xls";
        }

        readCSVFile("Alpha_Sprint_E.csv");//locOfCSV);

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


            for(int x = 1; x < listOfItems.size();x++) {
//                System.out.println(listOfItems.get(x).toString());
                Label labelType = new Label(0, x, listOfItems.get(x).getIssueType());
                Label labelIssue = new Label(1, x, listOfItems.get(x).getIssueKey());
                Label labelSummary = new Label(2, x, listOfItems.get(x).getIssueSummary());
                excelSheet.addCell(labelType);
                excelSheet.addCell(labelIssue);
                excelSheet.addCell(labelSummary);
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


    private static void createIndividualSheets(sprintObject obj, int count){
/*
        //Hyperinks https://stackoverflow.com/questions/16195140/how-do-i-activate-a-hyperlink-in-excel-after-writing-it-in-jexcel
 */
        String [] headers = {
                "Test Case Name:","Description: ", "Test Case Completed Date:",
                "Run By:","Start Date","Finish Date","Jira Ticket","Time(How long did it take","Environment",
                "Build #","Prerequisite","Os / Browser:", "Assumptions","Overall Pass or Fail"
        };

        String [] workHeaders = {
                "Steps #","Title","Action","Expected Result","Actual Results", "Pass / Fail", "Notes"
        };

        Label label;
        WritableSheet excelSheet = myFirstWbook.createSheet(obj.getIssueKey(),count);
        String baseJiraUrl = "https://medbridge.atlassian.net/browse/";
        String jiraLinkURL = baseJiraUrl+obj.getIssueKey();
        String linkDesc = obj.getIssueKey();
        WritableCellFormat header = new WritableCellFormat();

        try {

            header.setBackground(Colour.DARK_PURPLE);

            for (int x = 0; x < headers.length; x++) {
                label = new Label(0, x, headers[x]);
                excelSheet.addCell(label);

                //Test Case Name
                if(headers[x].equalsIgnoreCase("test case name")){
                    label = (new Label(1,x,obj.getIssueKey()));
                    excelSheet.addCell(label);
                }

                //Description
                if(headers[x].equalsIgnoreCase("description")){
                    label = (new Label(1,x,obj.getIssueSummary()));
                    excelSheet.addCell(label);
                }

                //Hyper Link
                if (headers[x].equalsIgnoreCase("jira ticket")) {
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

        }


    }

    private static void addHeaders(){
        String [] arrayHeaders ={"Issue Type","Jira Ticket", "Summary", "QA Owner","Story Points", "Result(P/F)","Notes"};
        WritableCellFormat header = new WritableCellFormat();

        WritableSheet excelSheet = myFirstWbook.getSheet(0);

        try {

            header.setBackground(Colour.DARK_PURPLE);

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
