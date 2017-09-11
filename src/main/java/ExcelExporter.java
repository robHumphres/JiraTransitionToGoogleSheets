
import jxl.Workbook;
import jxl.write.*;
import jxl.write.Number;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
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
            myFirstWbook.createSheet("Test Plans", 0);

            //addHeaders();

            WritableSheet excelSheet = myFirstWbook.getSheet(0);
            //Write them to test plan
            for(int x = 0; x < listOfItems.size();x++){
                System.out.println(listOfItems.get(x).toString());

                Label labelType = new Label(x + 1, 0, listOfItems.get(x).getIssueType());
                Label labelIssue = new Label(x + 1, x+1, listOfItems.get(x).getIssueKey());
                Label labelSummary = new Label(x + 1, x+2, listOfItems.get(x).getIssueSummary());
                excelSheet.addCell(labelType);
                excelSheet.addCell(labelIssue);
                excelSheet.addCell(labelSummary);

                //myFirstWbook.createSheet(listOfItems.get(x).getIssueKey(),x);
                myFirstWbook.write();
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

            myFirstWbook.write();

        }catch(Exception e){
            System.out.print("Error thrown while trying to do the headers. (addHeaders Method)" + e.getStackTrace());
        }

    }


}
