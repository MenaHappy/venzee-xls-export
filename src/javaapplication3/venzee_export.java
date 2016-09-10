
package javaapplication3;
import java.io.*;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;
import org.json.simple.parser.ParseException;

public class venzee_export {

	
	public void export(String path, String output) throws IOException, ParseException{
            System.out.println(path);
            //read Json File
            String filePath = path;
            
            JSONParser parser = new JSONParser();  
            JSONArray obj = (JSONArray) parser.parse(new FileReader(filePath));
            
            Workbook wb = new HSSFWorkbook();
            FileOutputStream fileOut;
                fileOut = new FileOutputStream(output);
                List<Sheet> sheets = new ArrayList<Sheet>();
           
           int i = 0;

           
            for (Object o : obj){
                JSONObject sheet = (JSONObject) o;
                Sheet temp = wb.createSheet((String) sheet.get("sheetName"));
                i = 0;
                JSONArray headers = (JSONArray) sheet.get("headers");
                String[] heads = new String[headers.size()];
                for(Object h : headers){
                    
                    heads[i] = h.toString();
                    i++;
                }
                Row row = temp.createRow((short)0);
                writeData(row, heads);
                
                JSONArray rows = (JSONArray) sheet.get("values");
                for(Object r : rows){
                    for(int j = 0; j < heads.length; j++){
                        System.out.println(r.toString());
                    }
                }
                
            }
            
            wb.write(fileOut);
            fileOut.close();
            
            
      
	   
	}

	private void writeData(Sheet sheet) {
		// TODO Auto-generated method stub
		Row row = sheet.createRow((short)0);
	    String[] data0 = {"date","itemID","location","typeID","CREATED_BY","CREATION_DATE","ENTITY_STATE","LAST_MODIFIED_BY","LAST_MODIFIED_DATE","qty"};
		writeData(row, data0);		
		writeDataSheet(sheet.createRow((short)  1 ),"3000003","2460");
		writeDataSheet(sheet.createRow((short)  2 ),"3000005","11934");
		writeDataSheet(sheet.createRow((short)  3 ),"3000009","0");
		writeDataSheet(sheet.createRow((short)  4 ),"3000012","5496");
		writeDataSheet(sheet.createRow((short)  5 ),"3000013","0");
		writeDataSheet(sheet.createRow((short)  6 ),"3000014","479");
		writeDataSheet(sheet.createRow((short)  7 ),"3000020","4339");
		writeDataSheet(sheet.createRow((short)  8 ),"3000021","0");
		writeDataSheet(sheet.createRow((short)  9 ),"3000022","17345");
		writeDataSheet(sheet.createRow((short)  10  ),"3000024","15174");
		writeDataSheet(sheet.createRow((short)  11  ),"3000027","0");
		writeDataSheet(sheet.createRow((short)  12  ),"3000041","0");
		writeDataSheet(sheet.createRow((short)  13  ),"3000042","6");
		writeDataSheet(sheet.createRow((short)  14  ),"3000047","0");
		writeDataSheet(sheet.createRow((short)  15  ),"3000054","0");
		writeDataSheet(sheet.createRow((short)  16  ),"3000055","1792");
		writeDataSheet(sheet.createRow((short)  17  ),"3000057","494");
		writeDataSheet(sheet.createRow((short)  18  ),"3000060","27");
		writeDataSheet(sheet.createRow((short)  19  ),"3000061","78");
		writeDataSheet(sheet.createRow((short)  20  ),"3000064","1");
		writeDataSheet(sheet.createRow((short)  21  ),"3000066","0");
		writeDataSheet(sheet.createRow((short)  22  ),"3000069","0");
		writeDataSheet(sheet.createRow((short)  23  ),"3000070","150");
		writeDataSheet(sheet.createRow((short)  24  ),"3000080","54");
		writeDataSheet(sheet.createRow((short)  25  ),"3000081","150");
		writeDataSheet(sheet.createRow((short)  26  ),"3000083","784");
		writeDataSheet(sheet.createRow((short)  27  ),"3000085","345");
		writeDataSheet(sheet.createRow((short)  28  ),"3000088","0");
		writeDataSheet(sheet.createRow((short)  29  ),"3000090","0");
		writeDataSheet(sheet.createRow((short)  30  ),"3000091","0");
		writeDataSheet(sheet.createRow((short)  31  ),"3000096","38");
		writeDataSheet(sheet.createRow((short)  32  ),"3000119","0");
		writeDataSheet(sheet.createRow((short)  33  ),"3000159","0");
		writeDataSheet(sheet.createRow((short)  34  ),"3000179","710");
		writeDataSheet(sheet.createRow((short)  35  ),"3000180","6407");
		writeDataSheet(sheet.createRow((short)  36  ),"3000181","3914");
		writeDataSheet(sheet.createRow((short)  37  ),"3000182","0");
		writeDataSheet(sheet.createRow((short)  38  ),"3000187","1563");
		writeDataSheet(sheet.createRow((short)  39  ),"3000189","0");
		writeDataSheet(sheet.createRow((short)  40  ),"3000199","1471");
		writeDataSheet(sheet.createRow((short)  41  ),"3000200","832");
		writeDataSheet(sheet.createRow((short)  42  ),"3000217","0");
		writeDataSheet(sheet.createRow((short)  43  ),"3000385","0");
		writeDataSheet(sheet.createRow((short)  44  ),"3000537","1343");
		writeDataSheet(sheet.createRow((short)  45  ),"3000538","41");
		writeDataSheet(sheet.createRow((short)  46  ),"3000539","1");
		writeDataSheet(sheet.createRow((short)  47  ),"3000737","836");
		writeDataSheet(sheet.createRow((short)  48  ),"3000877","7023");
		writeDataSheet(sheet.createRow((short)  49  ),"3000922","471");
		writeDataSheet(sheet.createRow((short)  50  ),"3000933","5318");
		writeDataSheet(sheet.createRow((short)  51  ),"3001268","12706");
		writeDataSheet(sheet.createRow((short)  52  ),"3001393","0");
		writeDataSheet(sheet.createRow((short)  53  ),"3001395","104");
		writeDataSheet(sheet.createRow((short)  54  ),"3001443","0");
		writeDataSheet(sheet.createRow((short)  55  ),"3001457","0");
		writeDataSheet(sheet.createRow((short)  56  ),"3001462","0");
		writeDataSheet(sheet.createRow((short)  57  ),"3002085","0");
		writeDataSheet(sheet.createRow((short)  58  ),"3002251","0");
		writeDataSheet(sheet.createRow((short)  59  ),"3002253","0");
		writeDataSheet(sheet.createRow((short)  60  ),"3002264","0");
		writeDataSheet(sheet.createRow((short)  61  ),"3003477","304");
		writeDataSheet(sheet.createRow((short)  62  ),"3003487","0");
		writeDataSheet(sheet.createRow((short)  63  ),"3003617","0");
		writeDataSheet(sheet.createRow((short)  64  ),"3003627","0");
		writeDataSheet(sheet.createRow((short)  65  ),"3003628","0");
		writeDataSheet(sheet.createRow((short)  66  ),"3003697","1537");
		writeDataSheet(sheet.createRow((short)  67  ),"3003707","0");
		writeDataSheet(sheet.createRow((short)  68  ),"3003911","2506");
		writeDataSheet(sheet.createRow((short)  69  ),"3004007","39");
		writeDataSheet(sheet.createRow((short)  70  ),"3004008","177");
		writeDataSheet(sheet.createRow((short)  71  ),"3004009","298");
		writeDataSheet(sheet.createRow((short)  72  ),"3004017","2282");
		writeDataSheet(sheet.createRow((short)  73  ),"3004021","0");
		writeDataSheet(sheet.createRow((short)  74  ),"3004127","0");
		writeDataSheet(sheet.createRow((short)  75  ),"3004168","0");
		writeDataSheet(sheet.createRow((short)  76  ),"3004213","147");
		writeDataSheet(sheet.createRow((short)  77  ),"3004239","1");
		writeDataSheet(sheet.createRow((short)  78  ),"3004240","0");
		writeDataSheet(sheet.createRow((short)  79  ),"3004242","341");
		writeDataSheet(sheet.createRow((short)  80  ),"3004243","251");
		writeDataSheet(sheet.createRow((short)  81  ),"3004434","0");
		writeDataSheet(sheet.createRow((short)  82  ),"3004467","268");
		writeDataSheet(sheet.createRow((short)  83  ),"3004511","0");
		writeDataSheet(sheet.createRow((short)  84  ),"3004737","0");
		writeDataSheet(sheet.createRow((short)  85  ),"3004750","0");
		writeDataSheet(sheet.createRow((short)  86  ),"3004847","539");
		writeDataSheet(sheet.createRow((short)  87  ),"3004895","149");
		writeDataSheet(sheet.createRow((short)  88  ),"3005293","2313");
		
		
	}

	private void writeDataSheet(Row row, String itemID, String k) {
		// TODO Auto-generated method stub
		row.createCell(0).setCellValue("08/01/2016 0:00:00");
		row.createCell(1).setCellValue(itemID);
		row.createCell(2).setCellValue("3150201000");
		row.createCell(3).setCellValue("1");
		row.createCell(4).setCellValue("");
		row.createCell(5).setCellValue("");
		row.createCell(6).setCellValue("ACTIVE");
		row.createCell(7).setCellValue("");
		row.createCell(8).setCellValue("");
		row.createCell(9).setCellValue(k);

	}

	private void writeDefinition(Sheet sheet) {
		
	    Row row = sheet.createRow((short)0);
	    String[] data0 = {"Id","date","itemID","location","typeID","CREATED_BY","CREATION_DATE","ENTITY_STATE","LAST_MODIFIED_BY","LAST_MODIFIED_DATE","qty"};
		writeData(row, data0);

		Row row1 = sheet.createRow((short)1);
	    String[] data1 = {"Display Name","date","itemID","location","typeID","CREATED_BY","CREATION_DATE","ENTITY_STATE","LAST_MODIFIED_BY","LAST_MODIFIED_DATE","qty"};
		writeData(row1, data1);
	
		Row row2 = sheet.createRow((short)2);
		String[] data2 = { "Datatype(string/date/numeric)","string","string","string","string","string","string","string","string","string","string"};
		writeData(row2, data2);

		Row row3 = sheet.createRow((short)3);
		String[] data3 = { "Cell Style","Sample","Sample","Sample","Sample","Sample","Sample","Sample","Sample","Sample","Sample"};
		writeData(row3, data3);
		
		Row row4 = sheet.createRow((short)4);
		String[] data4 = { "IsHidden flag","false","false","false","false","false","false","false","false","false","false"};
		writeData(row4, data4);
		
		Row row5 = sheet.createRow((short)5);
		String[] data5 = { "Multiple Occurrence?","false","false","false","false","false","false","false","false","false","false"};
		writeData(row5, data5);
		
		Row row6 = sheet.createRow((short)6);
		String[] data6 = { "Cell Style Override(style name)","","","","","","","","","",""};
		writeData(row6, data6);
		
		Row row7 = sheet.createRow((short)7);
		String[] data7 = { "Key fields(used for excel upload)","false","false","false","false","false","false","false","false","false","false"};
		writeData(row7, data7);
		
		Row row8 = sheet.createRow((short)8);
		String[] data8 = { "Group by field(used for excel upload)","false","false","false","false","false","false","false","false","false","false"};
		writeData(row8, data8);

	}

	private void writeData(Row row, String[] data1) {
		for (int i = 0; i < data1.length; i++) {
			if (data1[i].equals("false")){
				writeBooleanCell(row,i, false);
			} else {
				writeCell(row,i, data1[i]);
			}
			
		}
	}

	private void writeCell(Row row, int idx, String value) {
		row.createCell(idx).setCellValue(value);
	}
	
	private void writeBooleanCell(Row row, int idx, boolean value) {
		row.createCell(idx).setCellValue(value);
	}
	
}